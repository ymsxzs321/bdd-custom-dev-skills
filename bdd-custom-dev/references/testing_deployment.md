# 测试与部署实操经验（非 JSON 部分）

> 本文档记录 BddFuncs 二次开发中关于**编译环境、命令行、自测与 DLL 部署**的实战经验。
> JSON 解析相关的坑与经验单独归到业务实现中，这里不重复。

## 1. C# 语法版本兼容（编译环境）

- 本项目基于 **.NET Framework 4.x + 旧版 csproj**，自带的 `csc` 是 **C# 5** 编译器。
- **不支持**以下 C# 6+ 语法，写了会直接编译失败：
  - `is` 模式匹配：`if (obj is ArrayList arr)` → 须改为 `if (obj is ArrayList) { var arr = (ArrayList)obj; }`
  - 字符串插值 `$"..."` → 用 `string.Format` 或 `+` 拼接
  - 表达式体成员、null 条件 `?.` 之外的较新特性需谨慎
- **经验**：统一写 C# 5 兼容代码，确保 VS 旧 csproj 与命令行 `csc` 都能编译通过。

## 2. 在 PowerShell 中编译/运行外部命令（最稳姿势）

实战踩过的坑链：`/r:` 被 PowerShell 当成除法运算符 → 改用 `cmd /c` 后引号转义乱套 → 改用 `--%` 又导致 `>` 重定向失效 → `Get-Content` 读日志被安全策略拦截 → `Start-Process` 用相对路径时日志被写到意料之外的目录。

**经验（统一用这套）**：

```powershell
$dir = "d:\path\to\workdir"   # 工作目录（绝对路径）
Start-Process -FilePath "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" `
    -ArgumentList "/r:`"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Web.Extensions.dll`\"","Test.cs" `
    -WorkingDirectory $dir `
    -RedirectStandardOutput ($dir + "\build_out.log") `
    -RedirectStandardError  ($dir + "\build_err.log") `
    -Wait -NoNewWindow
```

要点：
- 日志路径**必须绝对路径**：`Start-Process` 的相对路径按调用方当前目录解析，会写到别处。
- 编译/运行后，用 `read_file` 工具读取 `.log` 看结果，**不要**依赖终端回显或 `Get-Content`（可能被安全策略或 CLIXML 包装干扰）。
- 用 `Test-Path "xxx.exe"` 判断是否生成成功；没生成就是编译失败，读 `build_err.log`。

## 3. 定位 msbuild（项目 Rebuild）

- `msbuild` 通常**不在全局 PATH**（Visual Studio 可能装在非 C 盘默认路径，如 `D:\Program Files\Microsoft Visual Studio\18\Community`）。
- 用 `vswhere` 精确定位：
  ```powershell
  & "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe" `
      -latest -products * -requires Microsoft.Component.MSBuild `
      -find "MSBuild\Current\Bin\MSBuild.exe"
  ```
- Rebuild（同样用 `Start-Process` 重定向日志到绝对路径）：
  ```powershell
  $msb  = "<上面定位到的 MSBuild.exe>"
  $proj = "d:\...\BddFuncs\BddFuncs.csproj"
  Start-Process -FilePath $msb -ArgumentList $proj, "/t:Rebuild",
      "/p:Configuration=Debug", "/p:Platform=AnyCPU", "/v:minimal" `
      -RedirectStandardOutput "d:\bdd_build.log" `
      -RedirectStandardError  "d:\bdd_build_err.log" -Wait -NoNewWindow
  ```
- 确认日志出现 `BddFuncs -> bin\Debug\BddFuncs.dll` 且无 error 即为成功。

## 4. 自测核心函数（与 Office/表堆堆解耦）

- 把要验证的核心逻辑（如某个业务方法）抽成**最小的临时 C# 程序**，用 `csc` 单独编译运行，**不要每次都 Build 整个 VSTO 插件**（慢、且依赖 Excel/Office）。
- 这正好满足 SKILL.md「函数测试与自我修复」的要求：只测核心函数、测完删测试代码。
- 自测文件用完即删，不残留到最终交付代码中（临时 `.cs` / `.exe` / `.log` 一并清理）。

## 5. 部署 DLL 到表堆堆安装目录

- 表堆堆是 Excel VSTO 插件，安装目录由注册表决定（详见 SKILL.md「函数开发完成后的构建与自动部署」）。
- 定位：`HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelReport` 的 `Manifest` 字段，取 `|` 前部分转本地路径的**目录**。
- 复制 `bin\Debug` 下**全部 `.dll`**（含生成的 DLL + 依赖 DLL），不要只复制主 DLL：
  ```powershell
  $manifest   = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelReport").Manifest
  $installDir = [System.IO.Path]::GetDirectoryName([System.Uri]::new($manifest.Split('|')[0]).LocalPath)
  Copy-Item "bin\Debug\*.dll" -Destination $installDir -Force
  ```
- **验证部署成功**：复制后检查目标目录中 `BddFuncs.dll` 的 `LastWriteTime` 与最新构建一致，即代表部署生效。
- 部署后重启 Excel/WPS 重新加载表堆堆，即可调用新接口函数。

## 6. msbuild Rebuild 步骤可能耗时较长、被自动跳过

- `msbuild /t:Rebuild` 是本项目里最慢的一步（要编译整个 VSTO 插件）。在某些执行环境下，若命令**约 10 秒内无返回**，会被自动判定为"耗时过长"而跳过（提示 `may take a long time ... skipped`），结果 `bin\Debug\BddFuncs.dll` 未更新、后续复制的是旧 DLL。
- 处理经验（二选一）：
  - **方式 A（推荐）：被跳过就重试**。用完全相同的 Rebuild 命令再跑一次，通常能正常完成并生成最新 DLL（实测：首次被跳过，重试后日志出现 `BddFuncs -> bin\Debug\BddFuncs.dll` 且无 error）。重试后务必核对日志确认有该输出行。
  - **方式 B：直接不用该步（仅限特定情况）**。若你**只改了纯 C# 业务逻辑**（如 `FuncsExe` 里的方法实现），且**没有新增引用、没有改动 COM 可见类型 / 接口签名**，可先不跑完整 Rebuild，而改用 §4 的独立 `csc` 自测验证核心逻辑，并确认 `bin\Debug` 中已有可运行的 DLL，再直接复制部署。注意：此方式部署的是**已有编译产物**，必须确保它确实包含你的改动，否则应回归方式 A 重新编译。
- 判断是否需要 Rebuild 的简易准则：只要动了 `InterFuncs.cs` 的接口/签名、新增了 `using` 引用或新增了依赖 DLL，就必须走方式 A 重新编译，不能跳过。
