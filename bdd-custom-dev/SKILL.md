---
name: 表堆堆二次开发
description: >
  This skill should be used when the user needs to develop custom functions for the 表堆堆 (Bdd) Excel/WPS add-in platform using C# and VSTO.
  It covers the BddFuncs secondary development template, including creating custom functions that appear in the function selection dialog,
  implementing interface functions callable via macros/VBA, integrating with 表堆堆's internal API (setVal, getVal, vars, etc.) — note that the reserved API is for selection-box functions only, not interface functions,
  and deploying custom DLLs. Trigger this skill when the user mentions 表堆堆二次开发, BddFuncs, extending 表堆堆 functionality,
  creating custom Excel functions for 表堆堆, or developing plugins using the 表堆堆 C# SDK.
---

# 表堆堆二次开发技能

## 概述

本技能用于指导基于 **表堆堆** (Bdd) 平台的二次开发。表堆堆是一个轻量级数据处理平台，以 Excel/WPS 插件形式呈现：用户把数据处理流程写成工作表上从上到下执行的一行行"函数 + 参数"（称为"堆表"），函数运行时在内存中生成变量（文本/数字/日期/工作簿/单元格/表格等）供后续引用。表堆堆提供基于 C# 语言的二次开发模板 **BddFuncs**，通过它开发的新函数可像内置函数一样集成到表堆堆中。

> 平台官网：https://sqlcel.com/ 　函数文档：https://sqlcel.com/funcs

## 何时使用本技能

- 用户需要为表堆堆开发自定义数据处理函数
- 用户需要在表堆堆的函数选择对话框中添加新函数
- 用户需要通过表堆堆宏函数调用自定义 C# 函数
- 用户需要集成外部 API（如 AI 识别、第三方服务）到表堆堆
- 用户询问 BddFuncs 模板的使用方法
- 用户需要了解选择框函数和接口函数的区别与选型

## 核心知识点

### 两种自定义函数

表堆堆二次开发支持两种自定义函数：

| 类型 | 调用方式 | 是否需要宏 | 适用场景 |
|------|---------|-----------|---------|
| **选择框函数** | 直接出现在函数选择对话框中 | 不需要 | 频繁通过界面调用的操作 |
| **接口函数** | 主要通过"宏"函数调用（宏 → VBA 桥接 → `f()` 接口） | **必须** | 后台自动调用、保持选择框整洁 |

两者通过同一个 **BddFuncs** 模板开发，功能能力相当。

> ⚠️ **开发前必须先确认函数类型**：选择框函数和接口函数是**两种完全不同的开发方式**（调用入口、部署形态、能否使用表堆堆预留 API 都不同）。**当用户没有明确说明要开发的是"选择框函数"还是"接口函数"时，必须先询问用户，让用户在二者之间做出选择，然后再开始开发。** 切勿自行默认其中一种。

### 开发前提

- **开发语言**：C#
- **技术栈**：Visual Studio Tools for Office (VSTO)
- **模板下载**：https://sqlcel.com/bdd/BddFuncs.zip
- **Visual Studio 配置**：安装时需勾选 Office/SharePoint 开发工作负载

## 适用范围与例外

> ⚠️ **基于 Visual Studio 的 Copilot 开发时，仅跳过 `testing_deployment.md` 及其系统/环境层面约束**：当用户是**在 Visual Studio 中借助其内置 Copilot（或等效的完整 VS 开发环境）来开发自定义函数**时，视为"在完整 IDE 环境中开发"，此时：
> - **不需要参考** `references/testing_deployment.md` 文件；
> - **不需要遵循**「经验沉淀（自动写入 testing_deployment.md）」小节，以及 `testing_deployment.md` 所记录的**系统/环境层面**约束（如命令行 `csc`/`msbuild` 姿势、`msbuild` 不在 PATH、Rebuild 被环境跳过、安装目录定位等）。
>
> **但以下测试与部署要求仍然必须遵循**（不因为是 VS Copilot 开发而豁免）：
> - 「函数测试与自我修复」中的**自测核心函数**（开发出的函数要自己测、测不过自己查因改正、测完删测试代码）；
> - 「函数开发完成后的构建与自动部署」与「部署方式」中的**自动重新生成（Rebuild）项目**与**自动复制全部 DLL 及其依赖到表堆堆安装目录**（确保用户马上可用）。
>
> 简言之：VS Copilot 开发只免去"参考 testing_deployment.md 及其中系统/环境坑"这一部分；核心自测与 DLL 自动部署照常。这些测试/部署约束仅在**非完整 VS IDE 环境**（例如借助外部 AI 助手、纯命令行、或缺少 VS 的编译/部署链路）开发时才适用。判断是否适用时，以"是否由 Visual Studio 的 Copilot 在 VS 内主导开发"为准。

## 开发工作流

### 场景一：开发选择框函数（显示在函数选择对话框中）

当需要让用户从表堆堆函数选择对话框直接调用新函数时，按以下步骤操作：

1. **定义新函数**：在 `newFuncs()` 方法中添加函数定义字典，包含 `funcName`（必填）、`keywords`、`funcCategory`、`description`、`param1`~`param12` 等属性
2. **实现执行逻辑**：在 `exefuncs(Excel.Range rn)` 的 switch 中添加 case 分支，通过 `rn.Offset[0, N]` 读取参数列的值
3. **（可选）实现辅助事件**：在 `alerts_selection_change` 和 `alerts_sht_change` 中实现参数辅助录入和数据校验
4. **编译部署（生成独立插件）**：构建 BddFuncs 项目，将其作为**独立的 Excel 插件（COM 加载项）**运行，与表堆堆**并行、相互独立**；也可封装为安装包分发给用户。注意：不是把 DLL 复制进表堆堆安装目录，而是让插件与表堆堆各自独立运行

### 场景二：开发接口函数（通过 f 接口 + VBA/宏调用）

当需要通过宏函数间接调用自定义逻辑时，按以下步骤操作：

1. **在 `f()` 接口中注册**：在 switch 中添加新函数的 case 分支，通过 `funcName` 区分
2. **编写实现方法**：在项目中编写具体业务逻辑（可放在 `FuncsExe` 类中，也可按习惯新建自己的类，并非必须放在 `FuncsExe` 类）
3. **选择调用方式（两种都经"宏"函数触发）**：
   - **方式 A · 经 VBA 桥接**：编写 VBA 过程，内部通过 `r.f("函数名", 参数...)` 调用接口函数，再由"宏"函数引用该 VBA 过程名
   - **方式 B · 宏直接调用（免 VBA）**：在堆表过程中直接用 `宏,f.函数名,返回值名称,参数1,参数2,...` 的格式调用，表堆堆会直接路由到 `f()` 接口，**无需编写 VBA 过程**
4. **在堆表过程中使用**：通过"宏"函数引用 VBA 过程名（方式 A），或直接用 `宏,f.函数名,...` 格式（方式 B）调用接口函数
5. **构建与自动部署（复制到表堆堆安装目录）**：开发完毕后需**自动重新生成项目**并（接口函数）**自动复制全部 DLL 及其依赖**到表堆堆安装目录。具体步骤见下方「函数开发完成后的构建与自动部署」一节。注意：接口函数依赖表堆堆内的宏来调用，因此必须把 DLL 放入表堆堆安装目录（与选择框函数的"独立插件"部署方式相反）
6. **交付说明（告知用户调用方式）**：完成部署后，AI 必须向用户说明以下信息（**按重要程度**）：

   - ⭐ **（最重要）函数名称**：即 `f()` 的 switch 中注册的 `funcName`
   - ⭐ **（最重要）参数名称**：每个参数的名称（及顺序）
   - **返回值的数据类型**：`f()` 返回值 `object` 实际指代的类型（如 `string` / `DataTable` / `double` 等）
   - **在表堆堆中的调用方式**（免 VBA）：
     ```
     宏,f.接口函数名称,返回值名称,参数1,参数2,...
     ```
     - `f.接口函数名称`：即 `funcName`
     - `返回值名称`：返回值写回表堆堆内存的变量名，**由用户自行命名、可自由修改**（后续引用保持一致即可）
     - `参数1,参数2,...`：依次对应 `f()` 的 `v2`、`v3`…… 业务参数

   > 也可通过"方式 A · 经 VBA 桥接"（先写 VBA 过程内部 `r.f("函数名", 参数...)`，再由"宏"函数引用该过程名）调用；但默认的、最直接的告知重点应放在上面的"宏,f.函数名,..."免 VBA 格式上。

### 函数测试与自我修复

> ⚠️ **开发出的函数必须自行测试；测试不通过要自己查因改正，不要抛给用户**：在**项目 Rebuild 之前先自己测试一遍**，只测**核心函数**，测完把测试代码删掉。

- **接口函数（测 `f()` 的分支核心函数）**：`f()` 的 switch 分支通常只是把参数转交给底层业务方法，例如 `return FuncsExe.twoNumsAdd(Convert.ToDouble(v2), Convert.ToDouble(v3));` —— 这种情况**只需单独测试 `FuncsExe.twoNumsAdd` 这个函数本身**（直接调用或用临时测试代码验证其输入/输出），不必走表堆堆整套调用。
- **选择框函数（测 `exefuncs` 的分支核心函数）**：直接测试 `exefuncs` 中对应 case 分支所调用的核心业务方法即可。

> ⚠️ **测试无需安装完整 VSTO 开发依赖**：自测核心函数时**不需要**为项目还原/安装整套 VSTO（Office 互操作、VSTO Runtime 等）开发依赖，也不必让整个 BddFuncs 插件项目能在 VS 里跑起来。**只要能把被测核心函数所在的 `.cs` 文件单独编译成 DLL 并运行即可**——把该函数（及其直接依赖的纯 C# 代码）抽到一个最小的临时项目/单文件里，用 `csc`（C# 编译器）编译为 DLL 或可执行文件运行验证。这样测试轻量、快速，且与表堆堆/Office 环境解耦。

**测试不通过的，AI 要自己排查原因并主动改正**：检查业务逻辑、参数类型转换（`Convert.ToDouble` / `ToString` 等）、空值处理等；**改完再测，直到通过**。不要只把问题或异常抛给用户就结束。

> **测试代码要删除**：用于自测的临时代码（临时 `Main`、测试按钮、临时调用语句等）在测试通过后必须删除，不要留在最终交付的代码里。

### 经验沉淀（自动写入 testing_deployment.md）

> ⚠️ **在测试和部署环节踩过的坑、以及获取的经验，必须自动沉淀到 `references/testing_deployment.md`**：每当你在测试或部署过程中**遇到新的坑、踩过的雷、或总结出更稳妥的做法**，都要**主动、自动地**把这条经验追加写入本项目技能下的 `references/testing_deployment.md`（即 `BddFuncs/skills/bdd-custom-dev/references/testing_deployment.md`）。不要等用户提醒，也不要只记在对话里就丢失。写入时遵循该文件现有结构（分节、要点、`PowerShell` 代码块等），保持与已有内容风格一致，避免重复。

> **区分"代码层面"与"系统/环境层面"（决定是否写入）**：
> - ✅ **要写入**：与**系统兼容性、DLL 编译/部署错误、命令行/工具链、注册表、Office/WPS 环境**等"系统/环境"层面相关的踩坑与经验。例如：PowerShell 下 `csc`/`msbuild` 的稳妥调用姿势、`.dll` 复制遗漏依赖、Rebuild 被环境跳过、安装目录定位、`msbuild` 不在 PATH 等。
> - ❌ **不要写入**：**C# 代码本身**层面的坑与经验，即纯业务逻辑、语法兼容（如 C# 5 限制）、类型转换、`Convert.ToDouble`/`ToString`、API 用法、空值处理等"写代码"相关的内容。此类经验应留在代码/对话中，不沉淀到 `testing_deployment.md`。

### 函数开发完成后的构建与自动部署

> ⚠️ **开发完毕后必须自动执行，不要留给用户手动操作**：函数开发（无论选择框函数还是接口函数）完成后，应立即**自动重新生成（Rebuild）项目**；若开发的是**接口函数**，还需**自动把生成的全部 DLL 及其依赖复制到表堆堆安装目录**，确保用户马上可用。

1. **自动重新生成项目**：对 BddFuncs 项目执行"重新生成"（Rebuild），使编译输出目录（`bin\Debug`）下的 DLL 保持最新。（选择框函数同样需要 Rebuild，但其产物以独立插件形态部署，见下方「部署方式」。）
2. **（仅接口函数）定位表堆堆安装目录**：表堆堆在注册表中的 ProgID 为 `ExcelReport`，其加载 DLL 的安装目录通过如下注册表项的 `Manifest` 值反查：
   - 注册表项：`HKCU\Software\Microsoft\Office\Excel\Addins\ExcelReport`
   - `Manifest` 值示例：`file:///D:/路径/ExcelReport/ExcelReport.vsto|vstolocal`
   - 取 `|` 之前的部分并转为本地路径，其**所在目录**（即 `.../路径/ExcelReport`）即为表堆堆安装目录。
3. **（仅接口函数）复制全部 DLL 及依赖**：将本项目编译输出（`bin\Debug`）下的**所有 `.dll` 文件（含生成的 DLL 及其全部依赖 DLL）**复制到上一步得到的安装目录。
   ```powershell
   # 第 2 步：读注册表定位表堆堆安装目录
   $manifest   = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelReport").Manifest
   $installDir = [System.IO.Path]::GetDirectoryName([System.Uri]::new($manifest.Split('|')[0]).LocalPath)
   # 第 3 步：复制本项目生成的所有 DLL（含依赖）到安装目录
   Copy-Item "bin\Debug\*.dll" -Destination $installDir -Force
   ```
4. **（仅接口函数）验证**：复制完成后，重启 Excel/WPS 或重新加载表堆堆，即可通过宏（直接 `f.函数名` 或经 VBA 桥接）调用新接口函数。

> 说明：复制 `bin\Debug\*.dll` 即可覆盖"生成的 DLL + 依赖 DLL"两类文件（依赖 DLL 在编译时已随 CopyLocal 复制到 `bin\Debug`）。若仅复制主 DLL 而遗漏依赖，运行时会因找不到依赖而加载失败。

### 与表堆堆内存交互（仅适用于选择框函数）

> ⚠️ **预留 API 仅供选择框函数使用，接口函数仅可使用 `f()`**。表堆堆为 BddFuncs 预留的 API（`setVal` / `getVal` / `vars` / `showParamInput` / `showFieldSlct` / `showShtInput`）**只能在选择框函数的执行上下文中调用**。**接口函数（经 `f()` 接口 + 宏/VBA 调用）只能使用 `f()` 这一个函数，其余所有预留 API 一律不得调用**；其实现只依赖普通 C# 业务逻辑，结果通过 `f()` 的返回值交由 VBA 处理。

在选择框函数的执行过程中，使用表堆堆预留的 API 与内存交互：

- `setVal(valName, val)` — 向表堆堆内存添加变量
- `getVal(valName)` — 获取内存中的变量
- `vars()` — 获取所有变量列表
- `showParamInput(dt, title)` — 弹出参数选择对话框
- `showFieldSlct(flds, title, multi)` — 弹出字段选择对话框
- `showShtInput(allShts, title)` — 弹出工作表名称录入对话框

## 部署方式

> 两种自定义函数的部署方式**不同**，关键区别如下：

### 选择框函数 → 生成独立插件（与表堆堆并行运行）

BddFuncs 模板本身就是一个 Excel 插件（VSTO/COM 加载项）项目。选择框函数开发完成后，作为**独立的 Excel 插件（COM 加载项）**与表堆堆**并行、相互独立**地运行（可单独启用/禁用），也可封装为安装包分发给用户。**不是把 DLL 复制进表堆堆的安装目录。**

### 接口函数 → 复制 DLL 到表堆堆安装目录

接口函数依赖表堆堆内的宏/VBA 桥接调用，需将构建生成的 DLL **直接复制到表堆堆的安装目录**，表堆堆在加载时自动加载 BddFuncs.dll，之后即可通过 `r.f()` 经宏/VBA 调用。**不是生成独立插件。**

> 安装目录定位与自动复制：开发完毕后**自动重新生成项目**，再通过注册表 `HKCU\Software\Microsoft\Office\Excel\Addins\ExcelReport` 的 `Manifest` 值反查表堆堆安装目录（`.vsto` 文件所在目录），将 `bin\Debug` 下**所有 `.dll`（含依赖）**复制过去。完整步骤与脚本见上方「函数开发完成后的构建与自动部署」。

### 封装安装包（适合分发）

将 BddFuncs 封装为安装程序，适合团队分发和版本升级（多用于选择框函数的独立插件形态）。

## 参考文档

详细的 API 接口定义、参数格式约定、VBA 调用示例和完整案例，请参阅：

- `references/api_reference.md` — 完整的 API 参考文档，包含 5 个关键接口函数、6 个辅助函数、Variable 类型结构、参数格式约定、部署细节和 VBA 调用模式
- `references/testing_deployment.md` — 测试与部署实操经验：C# 5 语法兼容、PowerShell 下用 csc/msbuild 编译运行的稳妥姿势、自测核心函数的方法、复制 DLL 到表堆堆安装目录的注意事项

## 注意事项

1. 选择框函数的 `exefuncs` 返回值为布尔型：`true` 表示成功，`false` 表示失败
2. 接口函数 `f()` 共 20 个参数：第 1 个 `funcName`（string，函数名）+ 第 2~20 个 `v2`~`v20`（共 19 个 object 业务入参）
3. **参数必填/选填判定**：参数**默认都是必填**（深绿色底色），只有在 `paramN` 说明文字里显式写了"（选填）"才变为选填（浅绿色底色）；判定只认"（选填）"三个字，不认语义。`【选项A/选项B】` 是**独立的另一维度**（限定取值范围），与必填/选填无关，可组合使用。界面底色仅为提示、不强制，必填参数仍需在 `exefuncs` 中自行校验是否为空。详见 `references/api_reference.md` §2.3
4. **预留 API 使用范围**：`setVal`/`getVal`/`vars`/`showParamInput`/`showFieldSlct`/`showShtInput` 等表堆堆预留 API **仅供选择框函数使用**；**接口函数（经 `f()` 接口）只能使用 `f()`，不可调用这些预留 API**。详见 `references/api_reference.md` §三、§四
5. 建议在 `exefuncs` 中妥善处理异常，避免中断整个堆表过程
6. 部分 WPS 版本对第三方 COM 加载项有限制，建议在 Excel 环境下开发并测试
7. 选择框函数沿用表堆堆堆表布局：A 列为函数名、B~M 列依次为参数 1~12；`exefuncs` 通过 `rn.Offset[0, N]` 读取对应参数
