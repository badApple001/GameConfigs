# GameConfigs 配置表导表工具

遍历 `Data/` 目录下所有 `.xlsx` 配置表，一键生成 **C# 配置类**（强类型、零反射）和 **bytes 二进制数据文件**，运行时通过 `GameSchemas` 静态类直接访问。

## 目录结构

```
GameConfigs/
├── Data/                  # 策划放的 xlsx 配置表(支持子目录)
├── gen_configs.py         # 导表脚本
├── gen_configs.json       # 生成配置(作用域/输出目录等)
└── README.md
```

产物输出（相对 Unity 工程根目录）：

| 产物 | 位置 |
|---|---|
| C# 代码 | `Assets/GameScripts/Generated/DataConfig/*.g.cs` |
| bytes 数据 | `Assets/GameArt/GameConfigBytes/*.bytes`（由 YooAsset 收集打包） |

## 环境依赖

- Python 3（加入 PATH，或安装 py 启动器）
- `pip install openpyxl`

## 如何导表

**方式一（推荐）**：Unity 编辑器菜单 `MyFrameworks/读表`
脚本输出回显到 Console，完成后自动刷新导入产物。

**方式二**：命令行

```bash
cd GameConfigs
python gen_configs.py
```

> 改完 Excel 保存后重新执行即可；生成前会自动清理旧的 `.g.cs` / `.bytes`。

## xlsx 格式约定

| 行 | 内容 | 说明 |
|---|---|---|
| 第1行 | 表注释名 | 取 A1；A1 为 `#`/空时取第一个有效列的第 1 行。仅用于类注释 |
| 第2行 | 字段名 | **为空或以 `#` 开头 => 该列忽略**（仅策划使用，不导出） |
| 第3行 | 字段类型 | 见下方类型表 |
| 第4行 | 字段注释 | 生成字段的 `/// <summary>` |
| 第5行起 | 数据 | 忽略列不导出；整行为空则跳过 |

示例：

| 关卡表 | | | |
|---|---|---|---|
| ID | LvPfbName | InitMoney | #备注 |
| int | string | int | string |
| 关卡ID | 关卡prefab名 | 初始钱数 | 这一列被忽略 |
| 1 | Lv_1 | 200 | 随便写 |

### 规则细节

- **第一个有效列（忽略列不算）必须为 `int`**，作为 ID 建立映射
- 支持的类型：`int` `long` `float` `double` `bool` `string` 及对应数组 `int[]` `float[]` 等
- **枚举类型**：类型写 `enum_枚举名`（如 `enum_ItemCategory`），详见下方「枚举类型」
- **数组单元格**：用 `,` `;` `|` 分隔均可（如 `1,2,3`、`0.5|1.5`）；单个数值视为单元素数组
- **空单元格**：数值 → 0，bool → false，string → `""`，数组 → 空数组，枚举 → 第一个值(0)
- **bool 单元格** 接受：`true/false/yes/no/y/n/1/0`（大小写不限）
- **公式单元格**取缓存值，请先在 Excel 中保存
- ID 重复时导表会打印警告，映射保留先出现的行
- 类型解析失败会报出具体位置（文件/表/行/字段）并中止
- `~$` 开头的 Excel 临时文件自动跳过；`#` 开头的 sheet 自动跳过
- 一个 xlsx 含多个 sheet 时，每个 sheet 生成一张表，表名 = `文件名_sheet名 + Config`（文件名已含 `Config` 后缀则不重复追加）

### 枚举类型

第 3 行类型写 `enum_枚举名`（`enum_` 前缀会被移除），脚本自动收集该列出现过的所有字符串，**按首次出现顺序从 0 开始编号**，在 `ConfigEnums.g.cs` 中生成枚举定义：

| ID | Category |
|---|---|
| int | enum_ItemCategory |
| 资源ID | 资源分类 |
| 100000 | Resource |
| 200000 | Tool |
| 300000 | Weapon |

生成：

```csharp
public enum ItemCategory
{
    Resource = 0,
    Tool = 1,
    Weapon = 2,
}
```

字段直接以枚举类型声明，使用侧：

```csharp
var item = GameSchemas.ItemDefConfig_0(300000);
if (item.Category == ItemCategory.Weapon) { ... }
```

规则细节：

- 枚举值就是 Excel 里的字符串，重复出现映射到同一个值
- **同名枚举跨列/跨表使用时值列表自动合并**（并集，仍按首次出现顺序），导表时会打印提示
- 也支持枚举数组：`enum_ItemCategory[]`，单元格写法同普通数组（`Weapon,Armor`）
- 空单元格 → 0（即第一个枚举值）
- 值文本会净化为合法 C# 标识符（如 `Magic Staff` → 成员 `Magic_Staff`，注释保留原始文本）
- 枚举在 bytes 中按 int32 存储值下标；**新增值会追加在末尾，不要删改已有值的顺序**（会影响已打包数据）

## gen_configs.json 配置项

| 键 | 默认值（当前项目） | 说明 |
|---|---|---|
| `namespace` | `MyFrameworks.DataConfig` | 生成 C# 类的作用域 |
| `schemasClassName` | `GameSchemas` | 总配置表类名 |
| `classSuffix` | `Config` | 表类名后缀（表名取自 xlsx 文件名） |
| `dataDir` | `Data` | xlsx 所在目录（相对本脚本） |
| `csharpOutDir` | `../Assets/GameScripts/Generated/DataConfig` | C# 输出目录 |
| `bytesOutDir` | `../Assets/GameArt/GameConfigBytes` | bytes 输出目录 |
| `bytesDirName` | `GameConfigBytes` | 默认文件加载时使用的目录名 |
| `bytesExtension` | `.bytes` | bytes 扩展名 |
| `multiSheet` | `true` | 多 sheet 每 sheet 一张表；`false` 只取第一个 sheet |
| `skipSheetPrefix` | `#` | 跳过以此前缀开头的 sheet |

## 运行时使用

```csharp
GameSchemas.Load();                            // 加载全部表(只需一次)
var row  = GameSchemas.LevelTableConfig(0);    // 按下标取一行
var row2 = GameSchemas.LevelTableConfig_0(1000); // 按ID取一行(第0列映射)
int n    = GameSchemas.LevelTableConfig_nums();  // 行数
```

每个配置类实现 `IConfigRow` 接口（`Id` 属性 + `Read(BinaryReader)`），字段即 xlsx 中的有效列。

### bytes 加载（YooAsset 接入案例）

`GameSchemas.BytesProvider` 是 bytes 加载入口（`Func<string, byte[]>`，参数为表名）。**本项目 bytes 走 YooAsset 打包、不进 StreamingAssets，且 YooAsset 未启用 Addressable——寻址必须用全路径**。完整接入示例（即本项目 `ConfigService` 的实现，`Assets/GameScripts/GamePlay/Service/ConfigService.cs`）：

```csharp
/// <summary>
/// 加载全部配置表（幂等，重复调用直接返回）
/// </summary>
public void LoadAll( )
{
    if ( IsLoaded ) return;

    //bytes 由 YooAsset 收集打包；未启用 Addressable，寻址用全路径 Assets/GameArt/GameConfigBytes/{表名}.bytes
    GameSchemas.BytesProvider = LoadBytesByYooAsset;
    GameSchemas.Load( );
    IsLoaded = true;
    CLog.Log( "[ConfigService] 配置表加载完成" );
}

private static byte[] LoadBytesByYooAsset( string tableName )
{
    string location = $"Assets/GameArt/GameConfigBytes/{tableName}.bytes"; 
    var handle = YooAssets.LoadRawFileSync( location );
    try
    {
        if ( handle.Status != EOperationStatus.Succeed )
        {
            CLog.LogError( "[ConfigService] 配置表加载失败: {0} | {1}", tableName, handle.LastError );
            return null;
        }
        return handle.GetRawFileData( );
    }
    finally
    {
        handle.Release( );
    }
}
```

`LoadAll` 在 `FunGame.Launch`（YooAsset 初始化完成后）调用，业务代码直接用 `GameSchemas.XXXConfig_0(id)` 即可，无需关心加载。

### 编辑器环境访问（免运行）

编辑器脚本（如多语言工具、自定义检视面板）不进入 Play 模式也能读配置，走 `Editor` 嵌套入口，API 与运行时一致：

```csharp
var item = GameSchemas.Editor.ItemDefConfig_0(100000);   // 按ID
var row  = GameSchemas.Editor.LevelTableConfig(0);       // 按下标
int n    = GameSchemas.Editor.MiscConfig_nums();         // 行数
```

- 首次访问自动用 `AssetDatabase.LoadAssetAtPath<TextAsset>` 从 `Assets/GameArt/GameConfigBytes/` 加载全部表（无需先设 `BytesProvider`，无需 Play 模式）
- 加载目录常量：`GameSchemas.Editor.EditorBytesDir`（由 `bytesOutDir` 自动推导生成）
- 整个 `Editor` 类包在 `#if UNITY_EDITOR` 内，**打包时完全剔除，不污染运行时代码**
- 编辑器下改过表并重新导表后，若需重新读取，重进 Play 或重编译触发重新加载即可（静态数据进程内只加载一次）

> 注意：`Editor` 入口仅供编辑器代码使用，运行时业务请用 `GameSchemas.XXXConfig_0(id)`（由 `ConfigService` 走 YooAsset 加载）。

## 常见问题

- **ID 取不到 / 抛 `KeyNotFoundException`**：检查表里是否有该 ID；注意 ID 列重复时只保留先出现的行（导表时有警告）。
- **报"未知类型"或"缺少类型"**：第 3 行类型只支持上表列出的写法，注意 `int[]` 的方括号是英文半角。
- **报"int类型得到非整数值"**：int 列填了小数；long 列同理。Excel 数值精度上限为 2^53。
- **新增表后 Unity 里找不到类**：确认导表执行成功，且 `Generated/DataConfig` 下有了新的 `.g.cs`。
- **热更/出包**：bytes 在 `Assets/GameArt` 下随 YooAsset 收集，改表后需重新构建资源。
