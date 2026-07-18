# 表堆堆二次开发 API 参考文档

## 一、BddFuncs 模板的 5 个关键接口函数

打开 BddFuncs 项目中的 `InterFuncs.cs` 文件，可以看到以下接口：

```csharp
public interface IInterFuncs
{
    bool exefuncs(Excel.Range rn);
    List<Dictionary<string, string>> newFuncs();
    void alerts_selection_change(Excel.Range rn);
    void alerts_sht_change(Excel.Range rn);
    object f(string funcName = null, object v2 = null, object v3 = null, object v4 = null, 
             object v5 = null, object v6 = null, object v7 = null, object v8 = null, 
             object v9 = null, object v10 = null, object v11 = null, object v12 = null, 
             object v13 = null, object v14 = null, object v15 = null, object v16 = null, 
             object v17 = null, object v18 = null, object v19 = null, object v20 = null);
}
```

### 接口函数详解

| 接口函数 | 调用时机 | 功能描述 |
|---------|---------|---------|
| `newFuncs()` | 表堆堆加载时 | 获取所有新开发的函数定义，自动注册到功能选择对话框目录树 |
| `exefuncs(Range rn)` | 用户点击"运行"时 | 执行具体的函数动作。`rn` 是调用该函数的 A 列单元格 |
| `alerts_selection_change(Range rn)` | 单元格选择改变时 | 用于实现参数辅助录入。`rn` 是选择发生改变的单元格 |
| `alerts_sht_change(Range rn)` | 单元格值改变时 | 用于数据校验和联动。`rn` 是值发生改变的单元格 |
| `f(funcName, v2...v20)` | 通过宏/VBA 调用时 | 预留接口函数，最多 20 个参数，通过 funcName 路由到不同功能 |

---

## 二、选择框函数开发详解

### 2.1 定义新函数（newFuncs）

> 新函数字典的 **key** 对应 BddFuncs 模板中 `Model.cs` 文件里的 `FuncModel` 模型属性。定义自定义函数时，key 应使用与 `FuncModel` 一致的属性名（如 `funcName`、`keywords`、`funcCategory`、`description`、`param1`~`param12`）。
>
> 选择框函数沿用表堆堆堆表布局：**A 列为函数名，B~M 列依次为参数 1~12**，`exefuncs` 通过 `rn.Offset[0, N]` 读取参数。

```csharp
public List<Dictionary<string,string>> newFuncs()
{
    var funcs = new List<Dictionary<string, string>>();
    funcs.Add(new Dictionary<string, string>
    {
        ["funcName"] = "工作表索引",
        ["keywords"] = "工作表名称，工作表位置",
        ["funcCategory"] = "自定义函数",
        ["description"] = "获取工作簿中所有工作表的位置和名称",
        ["param1"] = "给获取后生成的表取一个名称以便引用",
        ["param2"] = "工作簿名称"
    });
    return funcs;
}
```

### 2.2 函数属性说明

| 属性 | 是否必填 | 说明 |
|------|---------|------|
| `funcName` | **必填** | 函数名称，显示在函数选择对话框 A 列 |
| `keywords` | 选填 | 搜索关键字，多个用中文逗号分隔 |
| `funcCategory` | 选填 | 函数分类，对应目录树节点 |
| `description` | 选填 | 功能描述，显示在选择框右侧 |
| `param1` ~ `param12` | 选填 | 参数 1~12 的说明文字 |

### 2.3 参数的必填 / 选填约定

> **一句话规则：参数默认都是"必填"，只有在参数说明文字里写了 `（选填）` 才变成"选填"。**
> 判定依据完全来自 `param1`~`param12` 的**说明文字本身**（即你在 `newFuncs` 里写的字符串），表堆堆据此在界面上渲染不同底色。

#### 判定规则

| 参数说明文字 | 判定 | 界面底色 | 含义 |
|-------------|------|---------|------|
| 不含 `（选填）` 字样 | **必填项** | **深绿色** | 不填则函数无法正常执行 |
| 含 `（选填）` 字样 | **选填项** | **浅绿色** | 有默认值，可按需填写 |

#### 示例（文字 → 判定）

```csharp
["param1"] = "需读取的数据所在的单元格区域"
// 无"（选填）" → 必填项（深绿色）

["param2"] = "（选填）需读取的数据所在的单元格区域"
// 有"（选填）" → 选填项（浅绿色）

["param3"] = "单元格区域。如不填写则读取工作表中所有已用区域的数据"
// 虽然语义上"可不填"，但文字里没有"（选填）"两字 → 仍判定为必填项（深绿色）
```

> ⚠️ 判定只认 **`（选填）` 这三个字**，不认语义。想让某参数显示为选填，必须在说明文字中显式写上"（选填）"。

#### 与"选项枚举"格式的区别（两个独立概念，勿混淆）

`【选项A/选项B】` 格式与"必填/选填"**无关**，它是另一维度的约定——用于限定参数**只能在给定选项中取值**：

```csharp
["param4"] = "【输出值/输出公式】"
// 表示该参数只有"输出值""输出公式"两个可选值；
// 它本身是否必填，仍由是否含"（选填）"决定
```

可以两者组合，例如 `"（选填）【是/否】"` 表示：选填 + 取值限定为"是/否"。

#### 重要：界面底色只是"提示"，真正的校验要靠代码

表堆堆的深/浅绿色只是给使用者的**视觉提示**，它**不会强制**用户必填。因此在 `exefuncs` 中仍需自行校验必填参数是否为空，否则可能因取到 `null` 而报错：

```csharp
string filePath = pb.getstr(rn.Offset[0, 1]);
if (string.IsNullOrWhiteSpace(filePath))
{
    MessageBox.Show("参数1为必填项，请填写后再运行", "参数缺失");
    return false;   // 主动拦截，避免后续 NullReferenceException
}
```

### 2.4 实现函数执行动作（exefuncs）

表堆堆调用 `exefuncs` 时只传递调用该函数的 A 列单元格（`rn`）。通过 `rn.Offset[0, N]` 读取参数：

```csharp
public bool exefuncs(Excel.Range rn)
{
    string funcName = pb.getstr(rn, true);    // 获取 A 列函数名
    string param1 = pb.getstr(rn.Offset[0, 1]); // B 列 = 参数1
    string param2 = pb.getstr(rn.Offset[0, 2]); // C 列 = 参数2
    
    switch (funcName)
    {
        case "工作表索引":
            return FuncsExe.getSheetNames(rn);
        case "图片识别":
            return FuncsExe.RegnizePic(rn);
        default:
            MessageBox.Show("第" + rn.Row + "行的函数\"" + rn.Value + "\"无法识别", "无法识别的函数");
            return false;
    }
}
```

**返回值**：布尔型。`true` = 执行成功（触发字体倾斜、颜色变深红），`false` = 执行失败。

### 2.5 单元格选择改变事件（alerts_selection_change）

用于实现参数辅助录入。`rn` 是选择发生改变的单元格，`rn.Column` 判断当前在哪一列（2=B列/参数1，3=C列/参数2，以此类推）：

```csharp
public void alerts_selection_change(Excel.Range rn)
{
    Excel.Worksheet sht = rn.Parent;
    string funcName = pb.getstr(sht.Cells[rn.Row, 1], true);
    switch (funcName)
    {
        case "工作表索引":
            if (rn.Column == 3)  // C 列 = 参数2
            {
                getWks(rn, "请选择工作簿名称");
            }
            break;
        default:
            break;
    }
}
```

### 2.6 单元格值改变事件（alerts_sht_change）

当用户修改参数单元格值时触发，用于数据校验和联动逻辑：

```csharp
public void alerts_sht_change(Excel.Range rn)
{
    string targetValue = pb.getstr(rn);
    switch (targetValue)
    {
        case "工作表索引":
            if (rn.Column == 3)
            {
                // 值改变时的处理逻辑
            }
            break;
        default:
            break;
    }
}
```

---

## 三、接口函数开发详解

> ⚠️ **重要约束：接口函数仅可使用 `f()`，其余预留 API 一律不可用**
>
> 接口函数是**主要通过"宏"函数调用的**：在堆表过程中用"宏"函数 → 调用 VBA 过程 → 经 `f()` 接口间接执行自定义逻辑（精简版不支持宏，故接口函数不可用）。接口函数（即通过 `f()` 接口、经宏/VBA 间接调用的自定义函数）**只能使用 `f()` 这一个函数**，**不能使用**表堆堆为 BddFuncs 预留的其它辅助 API，包括：`setVal`、`getVal`、`vars`、`showParamInput`、`showFieldSlct`、`showShtInput`。
>
> 这些预留 API **仅供开发选择框函数使用**（在 `exefuncs` / `alerts_selection_change` / `alerts_sht_change` 等选择框函数生命周期内调用）。接口函数的实现应只依赖普通的 C# 业务逻辑（如文件读写、第三方 SDK 调用、返回结果给 VBA），不要在其中调用上述预留 API 与表堆堆内存交互。

### 3.1 f() 接口结构

`f()` 最多支持 20 个参数（v2~v20），第一个参数 `funcName`（string 类型）指定函数名，其余参数为 object 类型：

```csharp
public object f(
    string funcName = null, object v2 = null, object v3 = null, object v4 = null, 
    object v5 = null, ... , object v20 = null)
{
    switch (funcName)
    {
        case "twoNumsAdd":
            return FuncsExe.twoNumsAdd(Convert.ToDouble(v2), Convert.ToDouble(v3));
        case "图片识别":
            return FuncsExe.RegnizePic(v2.ToString());
        default:
            break;
    }
    return null;
}
```

### 3.2 接口函数的业务实现

不同于选择框函数的 `bool RegnizePic(Excel.Range rn)` 签名，接口函数的实现方法直接使用业务参数：

```csharp
public static string RegnizePic(string imagePath)
{
    var client = new Ocr(API_KEY, SECRET_KEY);
    client.Timeout = 60000;
    var imageBytes = File.ReadAllBytes(imagePath);
    var options = new Dictionary<string, object>
    {
        { "detect_direction", "true" },
        { "language_type", "ENG" }
    };
    var result = client.AccurateBasic(imageBytes, options);
    if (result["error_code"] == null)
    {
        string words = "";
        foreach (var word in result["words_result"])
            words += word["words"].ToString();
        return words;
    }
    return null;
}
```

---

## 四、表堆堆为 BddFuncs 准备的 6 个辅助接口函数

> ⚠️ **使用范围：仅限选择框函数**
>
> 本章节列出的 6 个辅助接口函数（`setVal`、`getVal`、`vars`、`showParamInput`、`showFieldSlct`、`showShtInput`）是表堆堆**预留给"选择框函数"**调用的——即只能在 `exefuncs`、`alerts_selection_change`、`alerts_sht_change` 等选择框函数执行上下文中使用。
>
> **接口函数（经 `f()` 接口 + 宏/VBA 调用）只能使用 `f()`，不可以使用这些预留 API。** 若接口函数需要"写入结果/变量"，应通过函数的返回值（`object`）把结果交回 VBA，由 VBA 侧再决定如何处理。详见 §三 的约束说明。

### 4.1 setVal — 向表堆堆内存添加变量

```csharp
void setVal(string valName, object val, string funcName = "", int excuteRow = 0, string DataType = "");
```

| 参数 | 必填 | 说明 |
|------|------|------|
| valName | **是** | 变量名称 |
| val | **是** | 变量值 |
| funcName | 否 | 生成这个变量的函数名 |
| excuteRow | 否 | 生成变量的函数所在工作表行号 |
| DataType | 否 | 数据类型（描述性分类，字符串）。常用取值：**文本、数字、日期、工作簿、单元格、表格、通用**；它对应 `val` 的实际 .NET 类型分别为 `string`、`数值`、`DateTime`、`Excel.Workbook`、`Excel.Range`、`DataTable`、`object`（详见 §4.4） |

```csharp
dynamic bdd = Globals.ThisAddIn.Application.COMAddIns.Item("ExcelReport").Object;
bdd.setVal("resultTable", dataTable, "自定义函数", currentRow, "表格");
```

### 4.2 getVal — 获取表堆堆内存变量

```csharp
object getVal(string valName);
```

```csharp
dynamic bdd = Globals.ThisAddIn.Application.COMAddIns.Item("ExcelReport").Object;
object val = bdd.getVal("tableName");
// 返回 Variable 类型对象，可通过 val.val 获取实际值
```

> **注意（源资料存在表述差异）**：官方教程对 `getVal` 返回值有两种说法——一说"返回该变量的 `val` 属性（即实际值）"，二说"返回 `Variable` 对象，需用 `val.val` 取值"。本技能按较新的详解版采用 **返回 `Variable` 对象** 的实现（即先取对象再访问 `.val` / `.DataType` / `.funcName` / `.excuteRow`）。若实际运行时 `val` 已是裸值，请直接使用 `val` 而无需再访问 `.val`。

### 4.3 vars — 获取所有内存变量列表

```csharp
List<string> vars();
```

```csharp
private void How_To_User_BDD_Vars()
{
    dynamic bdd = Globals.ThisAddIn.Application.COMAddIns.Item("ExcelReport").Object;
    List<string> vars = bdd.vars();
    foreach (string key in vars)
    {
        object val = bdd.getVal(key);
        if (val is Excel.Workbook)  // 判断数据类型
        {
            // 操作该变量
        }
    }
}
```

### 4.4 Variable 类型结构

```csharp
public class Variable
{
    public object val { get; set; }      // 变量值
    public string funcName { get; set; } // 生成这个变量的函数名
    public string valName { get; set; }  // 变量名称
    public string DataType { get; set; } // 数据类型
    public int excuteRow { get; set; }   // 生成变量函数所在工作表行号
}
```

> **`val` 的常用运行时类型（.NET 类型）**
>
> `val` 是 `object`，表堆堆实际存入的常见类型有下面四种，开发时按类型强转后使用：
>
> | 运行时类型 | 说明 | 典型来源 |
> |-----------|------|---------|
> | `System.Data.DataTable` | 表格数据 | 读取 Excel 工作表、`JSON转表格`、查询函数、`setVal(...,"表格")` |
> | `string` | 文本 | 录入文本、OCR 识别结果、`setVal(...,"文本")` |
> | `Excel.Range` | 单元格/区域引用 | 定义单元格、单元格函数、`setVal(...,"单元格")` |
> | `Excel.Workbook` | 工作簿对象 | 定义工作簿、`setVal(...,"工作簿")` |
>
> 取值时先用 `is` 判断再强转（注意 `Excel.Range`/`Excel.Workbook` 为 COM 互操作类型，需 `using Excel = Microsoft.Office.Interop.Excel;`）：
>
> ```csharp
> dynamic bdd = Globals.ThisAddIn.Application.COMAddIns.Item("ExcelReport").Object;
> var v = bdd.getVal("myVar");           // 返回 Variable 对象
> object raw = v.val;                     // 实际值
> switch (v.DataType)
> {
>     case "表格":
>         DataTable dt = raw as DataTable;
>         break;
>     case "文本":
>         string s = raw as string;
>         break;
>     case "单元格":
>         Excel.Range rng = raw as Excel.Range;
>         break;
>     case "工作簿":
>         Excel.Workbook wb = raw as Excel.Workbook;
>         break;
> }
> // 或直接按类型匹配
> if (raw is Excel.Workbook wb) { /* 操作工作簿 */ }
> ```
>
> 注：`DataType` 字符串（文本/数字/日期/工作簿/单元格/表格/通用）是**描述性分类**，与 `val` 的**实际 .NET 类型**是两回事——例如 `DataType="文本"` 时 `val` 是 `string`，`DataType="表格"` 时 `val` 是 `DataTable`。两者需对照使用。

### 4.5 弹框函数

| 函数 | 功能 | 参数说明 |
|------|------|---------|
| `showParamInput(DataTable dt, string title, bool self_design_visible = false)` | 弹出参数选择对话框 | dt=可选项表格，title=标题，self_design_visible=是否允许自定义参数 |
| `showFieldSlct(List<string> flds, string title = "", bool multi = true)` | 弹出字段选择对话框 | flds=可选项列表，title=标题，multi=是否允许多选 |
| `showShtInput(Dictionary<int, string> allShts, string title = "")` | 弹出工作表名称录入对话框 | allShts=工作表位置和名称字典，title=标题 |

---

## 五、VBA 调用模式

### 5.1 获取插件引用

```vba
Public Function r() As Object
    Set r = Application.COMAddIns("ExcelReport").Object
End Function

Public Function b() As Object
    Set b = Application.COMAddIns("BddFuncs").Object
End Function
```

### 5.2 通过表堆堆调用（DLL 需放入表堆堆安装目录）

```vba
Sub 通过表堆堆调用图片识别(imagePath As String)
    Dim rs As String
    rs = r.f("图片识别", imagePath)
    r.setval "rs", rs
End Sub
```

### 5.3 通过 BddFuncs 直接调用（需加载 BddFuncs COM 加载项）

```vba
Sub 通过BddFuncs调用图片识别(imagePath As String)
    Dim rs As String
    rs = b.f("图片识别", imagePath)
    r.setval "rs", rs
End Sub
```

### 5.4 在堆表过程中通过宏函数调用（两种格式）

**方式 A · 宏 → VBA 过程（桥接）**：先写好 VBA 过程，再由"宏"函数引用其过程名：

```
参数1：通过表堆堆调用图片识别（VBA 过程名）
参数3：D:\screenshots\captcha.png（传入 VBA 过程的参数值）
```

**方式 B · 宏直接调用 `f()`（免 VBA）**："宏"函数也可**不经过 VBA 桥接**直接调用接口函数，格式为：

```
宏,f.函数名,返回值名称,参数1,参数2,...
```

- `f.函数名`：固定前缀 `f.` + 你在 `f()` 的 switch 中注册的 `funcName`，表堆堆据此直接路由到 `f()` 接口
- `返回值名称`：接口函数 `f()` 的返回值写回表堆堆内存时使用的变量名（供后续函数引用），**该名称由用户自行命名、可自由修改**（只要后续引用处保持一致即可）
- `参数1,参数2,...`：依次传入 `f()` 的 `v2`、`v3`…… 业务参数

示例：

```
行号    函数              参数1~参数N
5      宏,f.图片识别,rs,D:\screenshots\captcha.png
```

表示：调用 `f("图片识别", "D:\screenshots\captcha.png")`，并把返回值存入变量 `rs`。该方式无需任何 VBA 代码。

---

## 六、部署方式

> **两种自定义函数的部署方式不同**：
> - **选择框函数** → 开发出的函数以**独立的 Excel 插件（COM 加载项）**形式与表堆堆**并行、相互独立**地运行（**不是**把 DLL 复制进表堆堆安装目录）。
> - **接口函数** → 将构建生成的 DLL **直接复制到表堆堆的安装目录**，表堆堆加载时自动加载 BddFuncs.dll，再通过宏/VBA 桥接调用（**不是**生成独立插件）。

### 选择框函数：生成独立插件（与表堆堆并行运行）

| 方式 | 操作步骤 | 优点 | 适用场景 |
|------|---------|------|---------|
| **COM 加载项（标准）** | 构建 BddFuncs 项目，作为独立 Excel 插件加载，与表堆堆并行运行 | 独立生命周期，可单独启用/禁用 | 正式使用、需要独立管理插件 |
| **封装安装包** | 将 BddFuncs 封装为安装程序 | 支持批量部署、版本升级 | 团队分发、正式发布 |

### 接口函数：复制 DLL 到表堆堆安装目录

| 方式 | 操作步骤 | 优点 | 适用场景 |
|------|---------|------|---------|
| **复制到表堆堆安装目录（标准）** | 将 `bin\Debug` 中生成的 **所有 DLL（含依赖 DLL）** 复制到表堆堆安装目录，表堆堆加载时自动加载 | 无需单独安装插件，依赖表堆堆内的宏/VBA 即可调用 | 接口函数的正式部署 |
| **（可选）封装安装包** | 与上同理，只是把 DLL 复制步骤打包为安装程序 | 支持批量分发 | 团队分发 |

> ⚠️ **开发完毕后自动执行**：接口函数开发完成后，应**自动重新生成（Rebuild）项目**，再将 `bin\Debug` 下**全部 `.dll`（生成的 DLL + 其依赖 DLL）**复制到表堆堆安装目录。仅复制主 DLL 而遗漏依赖会导致加载失败。

#### 如何定位表堆堆安装目录（读注册表）

表堆堆在注册表中的 ProgID 为 `ExcelReport`，其加载 DLL 的安装目录通过如下注册表项的 `Manifest` 值反查：

- 注册表项：`HKCU\Software\Microsoft\Office\Excel\Addins\ExcelReport`
- `Manifest` 值形如：`file:///D:/路径/ExcelReport/ExcelReport.vsto|vstolocal`
- 取 `|` 前部分转为本地路径，其**所在目录**（即 `.../路径/ExcelReport`）即为表堆堆安装目录。

PowerShell 示例：

```powershell
# 读注册表定位安装目录
$manifest   = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelReport").Manifest
$installDir = [System.IO.Path]::GetDirectoryName([System.Uri]::new($manifest.Split('|')[0]).LocalPath)
# 复制本项目生成的所有 DLL（含依赖）到安装目录
Copy-Item "bin\Debug\*.dll" -Destination $installDir -Force
```

---

## 七、两种自定义函数完整对比

| 对比维度 | 选择框函数 | 接口函数 |
|---------|-----------|---------|
| 调用方式 | 直接出现在函数选择对话框 | 通过"宏"函数调用（可直接 `f.函数名` 免 VBA，或经 VBA 桥接）→ f() 接口 |
| 是否需要宏 | 不需要 | **必须**支持宏 |
| 存储位置 | 函数选择对话框目录树 | 不占选择框空间 |
| 开发复杂度 | 稍高（需 newFuncs + exefuncs + 可选事件） | 较低（仅需在 f() 中加 case） |
| 精简版可用性 | ✅ 可用 | ❌ 不可用（精简版不支持宏） |
| 适用场景 | 频繁界面操作的函数 | 后台自动调用的操作 |
| 参数读取方式 | `rn.Offset[0, N]` 从单元格读取 | 直接通过 v2~v20 参数传入 |
| 能否使用表堆堆预留 API（setVal/getVal/vars/show*） | ✅ 可以 | ❌ 不可以（接口函数仅可使用 `f()`，其余预留 API 均不可用） |
| 部署方式 | 生成**独立插件**（COM 加载项），与表堆堆并行运行 | 将 DLL **复制到表堆堆安装目录**，由表堆堆自动加载 |

### 两者的联系

1. 通过同一个 **BddFuncs** 模板开发
2. 功能能力相当，可互相替代
3. 部署方式不同：**接口函数**的 DLL 放入表堆堆安装目录后由表堆堆自动加载；**选择框函数**作为独立插件（COM 加载项）加载，不与表堆堆安装目录耦合

> ⚠️ **常见误区纠正**：二者**并非都能调用预留 API**。只有**选择框函数**可以使用 `setVal` 等预留 API 与表堆堆内存交互；**接口函数**不得使用这些预留 API（详见 §三、§四 的约束说明）。接口函数如需把结果交给表堆堆，应通过 `f()` 的返回值传递给 VBA，再由 VBA 侧处理。

---

## 八、完整案例：图片文字识别

### 场景

在网页爬虫中使用"页面截图"截取验证码 → 自定义函数识别图片中的文字 → 录入网站。

### 选择框函数实现

#### 步骤 1：定义新函数

```csharp
funcs.Add(new Dictionary<string, string>
{
    ["funcName"] = "图片识别",
    ["keywords"] = "图片识别，图像识别，OCR",
    ["funcCategory"] = "自定义函数",
    ["description"] = "使用百度AI进行图片识别，将识别结果返回到Excel中",
    ["param1"] = "图片文件的完整路径"
});
```

#### 步骤 2：在 exefuncs 中添加分支

```csharp
case "图片识别":
    return FuncsExe.RegnizePic(rn);
```

#### 步骤 3：实现识别逻辑

```csharp
public static bool RegnizePic(Excel.Range rn)
{
    var client = new Ocr(API_KEY, SECRET_KEY);
    client.Timeout = 60000;
    string imagePath = pb.getstr(rn.Offset[0,1]);
    var imageBytes = File.ReadAllBytes(imagePath);
    var options = new Dictionary<string, object>
    {
        { "detect_direction", "true" },
        { "language_type", "ENG" }
    };
    var result = client.AccurateBasic(imageBytes, options);
    if (result["error_code"] == null)
    {
        string words = "";
        foreach (var word in result["words_result"])
            words += word["words"].ToString();
        rn.Offset[0,2].Value = words;   // 结果写入 C 列
        return true;
    }
    else
    {
        rn.Offset[0, 2].Value = false;
    }
    return false;
}
```

### 接口函数实现

#### 步骤 1：在 f() 中注册

```csharp
case "图片识别":
    return FuncsExe.RegnizePic(v2.ToString());
```

#### 步骤 2：实现识别逻辑（直接使用业务参数）

```csharp
public static string RegnizePic(string imagePath)
{
    var client = new Ocr(API_KEY, SECRET_KEY);
    client.Timeout = 60000;
    var imageBytes = File.ReadAllBytes(imagePath);
    var result = client.AccurateBasic(imageBytes, options);
    if (result["error_code"] == null)
    {
        string words = "";
        foreach (var word in result["words_result"])
            words += word["words"].ToString();
        return words;
    }
    return null;
}
```

### 在堆表流程中使用

```
行号    函数          参数1                    参数2      参数3
2      启动浏览器     myBrowser                12030
3      浏览网页       myBrowser                http://example.com/login
4      页面截图       myBrowser | //img[@id='captcha'] | D:\temp | captcha.png
5      宏            通过表堆堆调用图片识别      rs         =E4
6      录入文本       myBrowser | //input[@id='code'] | rs
7      页面点击       myBrowser | //button[@id='submit']
```
