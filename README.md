# bdd-custom-dev-skills

> 🛠️ 让 AI 帮你开发表堆堆的自定义函数 —— 用自然语言描述函数需求，AI 自动生成 C# 代码并编译为可部署的 Excel 插件

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel](https://img.shields.io/badge/Excel-支持-217346?logo=microsoft-excel)](https://sqlcel.com/)
[![WPS](https://img.shields.io/badge/WPS-支持-blue?logo=wps)](https://sqlcel.com/)
[![C#](https://img.shields.io/badge/C%23-二次开发-239120?logo=csharp)](https://sqlcel.com/)

---

## 📖 这是什么？

**表堆堆** 是一个以 Excel/WPS 插件形式呈现的轻量级数据处理平台。它提供了基于 C# + VSTO 的二次开发模板 **BddFuncs**，允许开发者扩展表堆堆的函数库。

**本仓库** 为 AI 编程助手（如 CodeBuddy、Cursor、Trae 等）提供了完整的表堆堆二次开发知识库，让 AI 能够：

- 🗣️ **理解** 用户用中文描述的函数需求
- ✍️ **生成** 完整的 C# 代码（实现 `IInterFuncs` 接口）
- 📋 **输出** 部署说明（编译、注册、分发）
- ✅ **提供** VBA 调用示例

> 💡 **一句话总结**：打开你的 AI 编程助手，说一句"写一个把 X小时X分 转成分钟的自定义函数"，AI 就能生成完整的 C# 代码、编译配置和部署指南。

---

## 🎯 适用场景

### 什么时候需要二次开发？

当表堆堆内置的 60+ 函数无法满足特定需求时，你可以通过二次开发来扩展：

| 用户说什么 | AI 会生成什么 |
|-----------|--------------|
| "写一个解析 PDF 发票内容的函数" | C# 代码 + PDF 解析库引用 + 部署说明 |
| "写一个调用企业微信 API 发消息的函数" | C# 代码 + HTTP 请求 + 鉴权逻辑 |
| "写一个把 JSON 转成表格的函数" | C# 代码 + Newtonsoft.Json 引用 |
| "写一个图片 OCR 识别的函数" | C# 代码 + 百度/阿里 OCR SDK 引用 |
| "写一个 AES 加解密的函数" | C# 代码 + 加密算法实现 |

### 两种函数类型

本仓库涵盖两种自定义函数开发模式：

| 类型 | 调用方式 | 是否需要宏 | 适用场景 |
|------|---------|-----------|---------|
| **选择框函数** | 直接出现在函数选择对话框 | 不需要 | 频繁手动调用的操作 |
| **接口函数** | 通过"宏"函数 + VBA 桥接调用 | **必须** | 后台自动化流程 |

AI 会根据用户需求自动推荐合适的类型，并生成对应的代码。

---

## 🧩 系统架构

1. **用户输入**：用自然语言描述函数需求
2. **AI 理解**：读取 SKILL.md + references/ 知识库
3. **类型确认**：AI 向用户确认选择框函数 or 接口函数
4. **代码生成**：生成完整的 C# 代码（实现 IInterFuncs 接口）
5. **部署输出**：提供编译、DLL 部署、COM 注册的完整说明
6. **完成**：自定义函数可在表堆堆中使用



## 📂 仓库结构

### 根目录文件

| 文件 | 说明 |
|------|------|
| `SKILL.md` | AI 核心指令 + 路由逻辑（AI 入口） |
| `README.md` | 本文件 |

### references/ 知识库目录

| 文件 | 说明 |
|------|------|
| `api_reference.md` | ⭐ 完整 API 参考（5 个接口函数 + 6 个辅助 API） |
| `bdd-custom-dev.md` | 二次开发详解（两种函数对比 + 完整案例） |
| `bdd-ai-api.md` | COM 接口（与表堆堆内存交互） |

## ⚠️ 常见陷阱（AI 必看）

本仓库的 `SKILL.md` 中内置了以下硬约束规则，确保 AI 生成的代码准确可用：

| # | 规则 |
|---|------|
| ⛔1 | **开发前必须先确认函数类型**：选择框函数 vs 接口函数，AI 必须先问用户 |
| ⛔2 | **pb.getstr() 的来源**：`pb` 是 BddFuncs 模板中预置的 `PublicFunc` 工具类实例 |
| ⛔3 | **参数 Offset 映射**：A 列是函数名，B~M 列是参数1~12，对应 `Offset[0, 1]` ~ `Offset[0, 12]` |
| ⛔4 | **接口函数不得使用预留 API**：`setVal`/`getVal`/`vars` 等仅供选择框函数使用 |
| ⛔5 | **部署方式区别**：选择框函数 → 独立 COM 加载项；接口函数 → DLL 复制到表堆堆安装目录 |

## 🔗 相关链接

- [表堆堆官网](https://sqlcel.com/)
- [表堆堆函数文档](https://sqlcel.com/funcs)
- [表堆堆下载](https://sqlcel.com/download/)
- [BddFuncs 二次开发模板下载](https://sqlcel.com/bdd/BddFuncs.zip)
- [CodeBuddy](https://codebuddy.ai/)

