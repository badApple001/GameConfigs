#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 配置表导表工具
=====================
遍历 Data 目录下所有 .xlsx, 生成 C# 配置类 + bytes 数据文件。

xlsx 格式约定:
    第1行: 表注释名 (取A1; A1为#或空时取第一个有效列的第1行), 仅用于类注释
    第2行: 字段名   (为空或以#开头 => 该列忽略, 仅策划使用)
    第3行: 字段类型 (int/long/float/double/bool/string 及对应数组 int[] 等)
    第4行: 字段注释 (用于字段的<summary>)
    第5行起: 数据   (忽略列不导出; 第一个有效列必须为int, 作为ID映射)

数组单元格用 , ; | 分隔均可, 空单元格 => 空数组/默认值。
公式单元格取其缓存值(请先在Excel中保存)。

用法: python gen_configs.py
配置: 同目录 gen_configs.json (不存在则自动生成默认配置)
"""
import json
import re
import struct
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("缺少依赖 openpyxl, 请先执行: pip install openpyxl")
    sys.exit(1)

# 兼容Windows GBK控制台
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "gen_configs.json"

DEFAULT_CONFIG = {
    "namespace": "FunGames.DataConfig",
    "schemasClassName": "GameSchemas",
    "classSuffix": "Config",
    "dataDir": "Data",
    "csharpOutDir": "../Assets/GameScripts/Generated/DataConfig",
    "bytesOutDir": "../Assets/StreamingAssets/GameConfigBytes",
    "bytesDirName": "GameConfigBytes",
    "bytesExtension": ".bytes",
    "multiSheet": True,
    "skipSheetPrefix": "#",
}

# 类型映射: 类型名 -> (C#类型, BinaryReader读方法, struct打包符, 数组元素?)
TYPE_MAP = {
    "int": ("int", "ReadInt32", "<i"),
    "long": ("long", "ReadInt64", "<q"),
    "float": ("float", "ReadSingle", "<f"),
    "double": ("double", "ReadDouble", "<d"),
    "bool": ("bool", "ReadBoolean", "<B"),
    "string": ("string", "ReadString", None),  # string 特殊处理(7bit长度前缀)
}

CS_KEYWORDS = {
    "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char",
    "checked", "class", "const", "continue", "decimal", "default", "delegate",
    "do", "double", "else", "enum", "event", "explicit", "extern", "false",
    "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit",
    "in", "int", "interface", "internal", "is", "lock", "long", "namespace",
    "new", "null", "object", "operator", "out", "override", "params", "private",
    "protected", "public", "readonly", "ref", "return", "sbyte", "sealed",
    "short", "sizeof", "stackalloc", "static", "string", "struct", "switch",
    "this", "throw", "true", "try", "typeof", "uint", "ulong", "unchecked",
    "unsafe", "ushort", "using", "virtual", "void", "volatile", "while",
}

ARRAY_SPLIT_RE = re.compile(r"[,;|]")


class GenError(Exception):
    """导表错误, 消息内需带表/行/列定位"""
    pass


# ---------------------------------------------------------------- 工具

def is_ignored_name(v) -> bool:
    """字段名为空或以#开头 => 忽略列"""
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s.startswith("#")


def sanitize_class_name(raw: str) -> str:
    parts = re.split(r"[^0-9A-Za-z_一-鿿]+", raw)
    name = "".join(p[:1].upper() + p[1:] for p in parts if p)
    if not name:
        raise GenError(f"无法从文件名生成类名: {raw}")
    if name[0].isdigit():
        name = "T" + name
    return name


def sanitize_field_name(raw: str) -> str:
    name = re.sub(r"[^0-9A-Za-z_一-鿿]", "_", raw.strip())
    if not name:
        raise GenError(f"字段名非法: {raw!r}")
    if name[0].isdigit():
        name = "_" + name
    if name in CS_KEYWORDS:
        name = "@" + name
    return name


# ---------------------------------------------------------------- 数据转换

def to_int(v, where):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return 0
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, (int, float)):
        f = float(v)
        if not f.is_integer():
            raise GenError(f"{where}: int类型得到非整数值 {v!r}")
        return int(f)
    s = str(v).strip()
    try:
        return int(s)
    except ValueError:
        try:
            f = float(s)
            if f.is_integer():
                return int(f)
        except ValueError:
            pass
    raise GenError(f"{where}: 无法解析为int: {v!r}")


def to_float(v, where):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        raise GenError(f"{where}: 无法解析为float: {v!r}")


def to_bool(v, where):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    s = str(v).strip().lower()
    if s in ("true", "yes", "y", "1"):
        return True
    if s in ("false", "no", "n", "0"):
        return False
    raise GenError(f"{where}: 无法解析为bool: {v!r}")


def to_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def convert_scalar(v, type_name, where):
    if type_name in ("int", "long"):
        return to_int(v, where)
    if type_name in ("float", "double"):
        return to_float(v, where)
    if type_name == "bool":
        return to_bool(v, where)
    if type_name == "string":
        return to_str(v)
    raise GenError(f"{where}: 未知类型 {type_name}")


def convert_array(v, elem_type, where):
    """数组单元格: 支持 , ; | 分隔; 空 => []; 单个数值 => 单元素数组"""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return []
    if isinstance(v, str):
        items = [s.strip() for s in ARRAY_SPLIT_RE.split(v)]
        items = [s for s in items if s != ""]
    else:
        items = [v]  # 单个数值视为单元素
    return [convert_scalar(it, elem_type, where) for it in items]


# ---------------------------------------------------------------- 二进制写入

def write_7bit_int(buf: bytearray, value: int):
    """与 C# BinaryWriter 的 7-bit encoded int 完全一致"""
    v = value & 0xFFFFFFFF
    while v >= 0x80:
        buf.append((v & 0x7F) | 0x80)
        v >>= 7
    buf.append(v)


def write_cs_string(buf: bytearray, s: str):
    """与 BinaryWriter.Write(string) 一致: 7bit字节数 + UTF8字节"""
    data = s.encode("utf-8")
    write_7bit_int(buf, len(data))
    buf.extend(data)


def write_value(buf: bytearray, v, type_name):
    if type_name == "string":
        write_cs_string(buf, v)
    elif type_name == "bool":
        buf.extend(struct.pack("<B", 1 if v else 0))
    else:
        buf.extend(struct.pack(TYPE_MAP[type_name][2], v))


def write_table_bytes(table, out_path: Path):
    buf = bytearray()
    buf.extend(struct.pack("<i", len(table["rows"])))
    for row in table["rows"]:
        for field, v in zip(table["fields"], row):
            t = field["type"]
            if t.endswith("[]"):
                elem = t[:-2]
                buf.extend(struct.pack("<i", len(v)))
                for item in v:
                    write_value(buf, item, elem)
            else:
                write_value(buf, v, t)
    out_path.write_bytes(bytes(buf))


# ---------------------------------------------------------------- 解析xlsx

def parse_sheet(ws, table_name, file_name):
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 4:
        return None  # 空sheet跳过

    r1 = rows[0]  # 表注释名行
    names = rows[1]  # 字段名行
    types = rows[2] if len(rows) > 2 else ()
    comments = rows[3] if len(rows) > 3 else ()

    # 有效列下标
    col_idx = [i for i in range(len(names)) if not is_ignored_name(names[i])]
    if not col_idx:
        return None

    ctx = f"{file_name} -> {table_name}"

    # 表注释名: A1有效则用A1, 否则取第一个有效列的第1行
    title = None
    if r1 and not is_ignored_name(r1[0]):
        title = str(r1[0]).strip()
    elif col_idx and r1 and len(r1) > col_idx[0] and not is_ignored_name(r1[col_idx[0]]):
        title = str(r1[col_idx[0]]).strip()
    else:
        title = table_name

    fields = []
    seen = set()
    for i in col_idx:
        fname = sanitize_field_name(str(names[i]).strip())
        if fname in seen:
            raise GenError(f"{ctx}: 字段名重复 {fname}")
        seen.add(fname)
        tname = str(types[i]).strip() if i < len(types) and types[i] is not None else ""
        if not tname:
            raise GenError(f"{ctx}: 字段 {fname} 缺少类型(第3行)")
        base = tname[:-2] if tname.endswith("[]") else tname
        if base not in TYPE_MAP:
            raise GenError(f"{ctx}: 字段 {fname} 未知类型 {tname}")
        comment = ""
        if i < len(comments) and comments[i] is not None:
            comment = str(comments[i]).strip()
        fields.append({"name": fname, "type": tname, "comment": comment or fname, "col": i})

    # 第一个有效列必须为int
    if fields[0]["type"] != "int":
        raise GenError(f"{ctx}: 第一个有效列 {fields[0]['name']} 必须为int类型(作为ID), 当前为 {fields[0]['type']}")

    # 数据行
    data_rows = []
    id_seen = {}
    for r_idx, row in enumerate(rows[4:], start=5):
        vals = [row[i] if i < len(row) else None for i in col_idx]
        if all(v is None or (isinstance(v, str) and v.strip() == "") for v in vals):
            continue  # 跳过空行
        out = []
        for field, v in zip(fields, vals):
            where = f"{ctx} 第{r_idx}行[{field['name']}]"
            t = field["type"]
            if t.endswith("[]"):
                out.append(convert_array(v, t[:-2], where))
            else:
                out.append(convert_scalar(v, t, where))
        row_id = out[0]
        if row_id in id_seen:
            print(f"  [警告] {ctx}: ID {row_id} 重复(第{id_seen[row_id]}行与第{r_idx}行), 映射将保留先出现的行")
        else:
            id_seen[row_id] = r_idx
        data_rows.append(out)

    return {"name": table_name, "title": title, "fields": fields, "rows": data_rows, "source": file_name}


def collect_tables(cfg):
    data_dir = (SCRIPT_DIR / cfg["dataDir"]).resolve()
    if not data_dir.is_dir():
        raise GenError(f"Data目录不存在: {data_dir}")

    xlsx_files = sorted(p for p in data_dir.rglob("*.xlsx") if not p.name.startswith("~$"))
    if not xlsx_files:
        raise GenError(f"{data_dir} 下没有找到任何 .xlsx 文件")

    tables = []
    used_names = {}
    for fp in xlsx_files:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        try:
            sheets = [ws for ws in wb.worksheets if not ws.title.startswith(cfg["skipSheetPrefix"])]
            if not cfg["multiSheet"]:
                sheets = sheets[:1]
            for ws in sheets:
                base = sanitize_class_name(fp.stem)
                if len(sheets) > 1:
                    base += "_" + sanitize_class_name(ws.title)
                table_name = base + cfg["classSuffix"]
                if table_name in used_names:
                    raise GenError(f"表名冲突: {table_name} ({fp.name} 与 {used_names[table_name]})")
                t = parse_sheet(ws, table_name, fp.name)
                if t is None:
                    print(f"  [跳过] {fp.name} -> {ws.title} (无有效内容)")
                    continue
                used_names[table_name] = fp.name
                tables.append(t)
        finally:
            wb.close()
    return tables


# ---------------------------------------------------------------- C# 代码生成

CS_HEADER = "// <auto-generated> 由 gen_configs.py 生成, 请勿手改。"


def xml_comment(text, indent):
    pad = " " * indent
    return f"{pad}/// <summary>\n{pad}/// {text}\n{pad}/// </summary>\n"


def gen_table_cs(table, ns):
    fields = table["fields"]
    has_array = any(f["type"].endswith("[]") for f in fields)
    id_field = fields[0]["name"]

    sb = [CS_HEADER + f" 源表: {table['source']}\n"]
    sb.append("using System.IO;\n\n")
    sb.append(f"namespace {ns}\n{{\n")
    sb.append(xml_comment(table["title"], 4))
    sb.append(f"    public class {table['name']} : IConfigRow\n    {{\n")
    for f in fields:
        cs_type = TYPE_MAP[f["type"][:-2]][0] + "[]" if f["type"].endswith("[]") else TYPE_MAP[f["type"]][0]
        sb.append(xml_comment(f["comment"], 8))
        sb.append(f"        public {cs_type} {f['name']};\n")
    sb.append(f"\n        public int Id => {id_field};\n")
    # Read
    sb.append("\n        public void Read(BinaryReader reader)\n        {\n")
    if has_array:
        sb.append("            int __c;\n")
    for f in fields:
        t = f["type"]
        name = f["name"]
        if t.endswith("[]"):
            elem = t[:-2]
            reader_fn = TYPE_MAP[elem][1]
            cs_elem = TYPE_MAP[elem][0]
            sb.append(f"            __c = reader.ReadInt32();\n")
            sb.append(f"            {name} = new {cs_elem}[__c];\n")
            sb.append(f"            for (int __i = 0; __i < __c; __i++) {name}[__i] = reader.{reader_fn}();\n")
        else:
            sb.append(f"            {name} = reader.{TYPE_MAP[t][1]}();\n")
    sb.append("        }\n")
    sb.append("    }\n}\n")
    return "".join(sb)


def gen_base_cs(ns):
    return (
        CS_HEADER + "\n"
        "using System.IO;\n\n"
        f"namespace {ns}\n{{\n"
        "    /// <summary>\n"
        "    /// 配置行解析接口(所有配置类实现)\n"
        "    /// </summary>\n"
        "    public interface IConfigRow\n"
        "    {\n"
        "        /// <summary>第一列ID</summary>\n"
        "        int Id { get; }\n"
        "        /// <summary>从二进制流解析一行数据</summary>\n"
        "        void Read(BinaryReader reader);\n"
        "    }\n"
        "}\n"
    )


def camel(s):
    return s[:1].lower() + s[1:]


def gen_schemas_cs(tables, ns, cls_name, cfg):
    sb = [CS_HEADER + "\n"]
    sb.append("using System;\nusing System.Collections.Generic;\nusing System.IO;\n\n")
    sb.append(f"namespace {ns}\n{{\n")
    sb.append("    /// <summary>\n")
    sb.append("    /// 总配置表: Load()加载全部配置; XXXConfig(下标); XXXConfig_0(ID); XXXConfig_nums(行数)\n")
    sb.append("    /// </summary>\n")
    sb.append(f"    public static class {cls_name}\n    {{\n")
    sb.append("        /// <summary>自定义bytes加载(优先级最高), 参数为表名, 返回文件字节。Android等平台请设置它。</summary>\n")
    sb.append("        public static Func<string, byte[]> BytesProvider;\n\n")
    sb.append(f"        public const string BytesDirName = \"{cfg['bytesDirName']}\";\n")
    sb.append(f"        public const string BytesExtension = \"{cfg['bytesExtension']}\";\n\n")
    sb.append("        public static bool IsLoaded { get; private set; }\n\n")

    # 存储字段
    for t in tables:
        sb.append(f"        private static List<{t['name']}> _{camel(t['name'])}List;\n")
        sb.append(f"        private static Dictionary<int, {t['name']}> _{camel(t['name'])}Map;\n")
    sb.append("\n")

    # Load
    sb.append("        /// <summary>加载所有配置表</summary>\n")
    sb.append("        public static void Load()\n        {\n")
    for t in tables:
        sb.append(f"            Load{t['name']}();\n")
    sb.append("            IsLoaded = true;\n        }\n\n")

    # 每表加载 + 3个访问方法
    for t in tables:
        n = t["name"]
        lst = f"_{camel(n)}List"
        mp = f"_{camel(n)}Map"
        sb.append(xml_comment(f"加载 {t['title']}({n})", 8))
        sb.append(f"        private static void Load{n}()\n        {{\n")
        sb.append(f"            byte[] bytes = LoadBytes(\"{n}\");\n")
        sb.append("            using (var reader = new BinaryReader(new MemoryStream(bytes)))\n            {\n")
        sb.append("                int count = reader.ReadInt32();\n")
        sb.append(f"                {lst} = new List<{n}>(count);\n")
        sb.append(f"                {mp} = new Dictionary<int, {n}>(count);\n")
        sb.append("                for (int i = 0; i < count; i++)\n                {\n")
        sb.append(f"                    var row = new {n}();\n")
        sb.append("                    row.Read(reader);\n")
        sb.append(f"                    {lst}.Add(row);\n")
        sb.append(f"                    if (!{mp}.ContainsKey(row.Id)) {mp}.Add(row.Id, row);\n")
        sb.append("                }\n            }\n        }\n\n")
        sb.append(xml_comment(f"{t['title']} - 按下标获取", 8))
        sb.append(f"        public static {n} {n}(int index) => {lst}[index];\n\n")
        sb.append(xml_comment(f"{t['title']} - 按ID(第0列)获取", 8))
        sb.append(f"        public static {n} {n}_0(int id)\n        {{\n")
        sb.append(f"            if ({mp}.TryGetValue(id, out var row)) return row;\n")
        sb.append(f"            throw new KeyNotFoundException(\"{n} 不存在ID: \" + id);\n")
        sb.append("        }\n\n")
        sb.append(xml_comment(f"{t['title']} - 行数", 8))
        sb.append(f"        public static int {n}_nums() => {lst}.Count;\n\n")

    # LoadBytes
    sb.append("        private static byte[] LoadBytes(string tableName)\n        {\n")
    sb.append("            if (BytesProvider != null) return BytesProvider(tableName);\n")
    sb.append("#if UNITY_5_3_OR_NEWER\n")
    sb.append("            // 注意: Android上StreamingAssets在jar内, File不可读, 请设置 BytesProvider\n")
    sb.append("            string path = Path.Combine(UnityEngine.Application.streamingAssetsPath, BytesDirName, tableName + BytesExtension);\n")
    sb.append("#else\n")
    sb.append("            string path = Path.Combine(AppContext.BaseDirectory, BytesDirName, tableName + BytesExtension);\n")
    sb.append("#endif\n")
    sb.append("            return File.ReadAllBytes(path);\n")
    sb.append("        }\n")
    sb.append("    }\n}\n")
    return "".join(sb)


# ---------------------------------------------------------------- 主流程

def clean_dir(path: Path, pattern: str):
    if path.exists():
        for f in path.glob(pattern):
            f.unlink()


def main():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=4), encoding="utf-8")
        print(f"已生成默认配置: {CONFIG_PATH.name}")
    cfg = dict(DEFAULT_CONFIG)
    cfg.update({k: v for k, v in json.loads(CONFIG_PATH.read_text(encoding="utf-8")).items() if not k.startswith("_")})

    print(f"[1/3] 扫描 {cfg['dataDir']} ...")
    tables = collect_tables(cfg)
    if not tables:
        raise GenError("没有解析出任何配置表")

    cs_dir = (SCRIPT_DIR / cfg["csharpOutDir"]).resolve()
    bytes_dir = (SCRIPT_DIR / cfg["bytesOutDir"]).resolve()
    cs_dir.mkdir(parents=True, exist_ok=True)
    bytes_dir.mkdir(parents=True, exist_ok=True)
    clean_dir(cs_dir, "*.g.cs")
    clean_dir(bytes_dir, "*" + cfg["bytesExtension"])

    print(f"[2/3] 生成 C# -> {cs_dir}")
    ns = cfg["namespace"]
    (cs_dir / "ConfigBase.g.cs").write_text(gen_base_cs(ns), encoding="utf-8-sig")
    (cs_dir / f"{cfg['schemasClassName']}.g.cs").write_text(
        gen_schemas_cs(tables, ns, cfg["schemasClassName"], cfg), encoding="utf-8-sig")
    for t in tables:
        (cs_dir / f"{t['name']}.g.cs").write_text(gen_table_cs(t, ns), encoding="utf-8-sig")

    print(f"[3/3] 生成 bytes -> {bytes_dir}")
    for t in tables:
        out = bytes_dir / f"{t['name']}{cfg['bytesExtension']}"
        write_table_bytes(t, out)
        print(f"  √ {t['name']:<24} {len(t['rows']):>4} 行  {out.stat().st_size:>7} B  ({t['title']})")

    print(f"\n完成! 共 {len(tables)} 张表。")
    print(f"调用方式: {cfg['schemasClassName']}.Load();  {cfg['schemasClassName']}.{tables[0]['name']}_0(1);")


if __name__ == "__main__":
    try:
        main()
    except GenError as e:
        print(f"\n[导表失败] {e}")
        sys.exit(1)
