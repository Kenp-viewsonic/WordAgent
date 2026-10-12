<div align="center">
  <img src="apps/word-addin/public/icon-80.png" width="96" alt="Word Agent" />
  <h1>Word Agent</h1>
  <p><b>Windows 桌面 Word 本地智能体</b> — 让 AI 直接读懂、改写、整理你的文档</p>
  <p>
    <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-0078D6?logo=windows11" alt="platform" />
    <img src="https://img.shields.io/badge/Microsoft%20Word-2016%2B-2B579A?logo=microsoftword" alt="word" />
    <img src="https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?logo=dotnet" alt="dotnet" />
    <img src="https://img.shields.io/badge/Node.js-SEA%20打包-339933?logo=nodedotjs" alt="node" />
    <img src="https://img.shields.io/badge/阶段-Beta%20可分发测试-orange" alt="stage" />
    <img src="https://img.shields.io/badge/license-TBD-lightgrey" alt="license" />
  </p>
</div>

---

## ✨ 为什么选 Word Agent

|  |  |
|---|---|
| 🚀 **一键安装，随 Word 启动** | 纯 COM 加载项注册（非 Web Add-in 侧载），`Install.exe` 装完重启 Word 即出按钮 —— 无信任中心配置、无 Node 环境、无黑色命令行窗口 |
| 🤖 **真正的文档操作** | 24 个文档工具：读写段落、查找替换、样式排版、表格增删改、流程图插入……全部运行在 Word 对象模型上，带撤销合并（一次 Ctrl+Z 回滚一整步） |
| 🛩️ **Autopilot 多步执行** | 勾选 Autopilot 后连贯执行多轮任务不打断；每次发送自动建立 ⏪ 检查点，一键撤销本轮文档改动 / 回到发送前 |
| 📚 **本地知识库** | 上传文档建立检索库，回答带来源引用；支持存档导出 / 导入 |
| 🔒 **本地优先，安全加固** | 服务只绑定 `127.0.0.1`，全部 API 要求随机 token 鉴权，CORS 白名单 —— 恶意网页无法读取你的知识库与配置 |
| 📊 **用量与账户面板** | Token 消耗、工具调用排行、失败率一目了然；模型配置（BYOK）与预设方案集中管理 |
| 🧩 **自带你的模型** | OpenAI 兼容接口任意接入：OpenAI / DeepSeek / 通义 / Kimi / Ollama 本地模型……API Key 只存本机 |

## 🚀 快速开始

**普通用户**

1. 下载 [`OfficeAgent-Setup-<版本>.exe`](dist-setup/)（发布页获取）
2. 双击安装（可选"所有用户 / 仅当前用户"，无需管理员亦可）
3. 打开 Word → 「开始」选项卡 → **Word Agent → 打开助手**
4. 在面板右上角 ⚙️ 配置模型供应商（Base URL + API Key + 模型）→ 开始对话

> 安装/卸载前请完全退出 Word；首次使用需安装 [WebView2 运行时](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)（安装器会自动检测并引导，Win10/11 一般自带）。

**开发者**

```bash
npm install               # 安装依赖
npm run dev:agent         # 本地 Agent：http://127.0.0.1:8787
npm run dev:addin         # UI 开发服务器：http://localhost:3001

npm run package           # 构建 dist/ 并编译安装器 → dist-setup/
npm run smoke             # 安全冒烟（token 鉴权 / CORS / 启动）
npm run test:tools        # 24 文档工具回归（真实 Word 实例，34 用例）
```

COM 加载项开发调试：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register-com.ps1    # 注册（免管理员）
powershell -ExecutionPolicy Bypass -File scripts\unregister-com.ps1  # 注销
```

## 🏗️ 架构

```mermaid
graph TB
    subgraph WINWORD.EXE
        COM["COM 薄壳加载项<br/>Ribbon 按钮 + 任务窗格<br/>LoadBehavior=3 自动加载"]
        OM["Word 对象模型执行器<br/>24 个文档工具 + 撤销合并"]
    end
    subgraph TRAY["托盘常驻"]
        API["Agent API (127.0.0.1:8787)<br/>聊天 / Agent Loop / 知识库 / 用量统计<br/>token 鉴权 + CORS 白名单"]
    end
    UI["TS UI（聊天 / 知识库 / 检查点 / 用量）<br/>WebView2 加载"]
    COM -- Ribbon 点击 --> UI
    UI -- "fetch + SSE（X-Agent-Token）" --> API
    UI -- "文档工具 ⇄ WebView2 桥" --> OM
    COM -. "HKCU Run 开机自启" .- TRAY
    MODEL["任意 OpenAI 兼容模型"] <-- API
```

- **agent.exe**：esbuild + Node SEA 打包的独立服务，零 Node 依赖
- **OfficeAgent.Tray.exe**：C# WinForms 托盘壳（单实例 / 自启 / 服务守护 / 端口处置）
- **OfficeAgent.ComHost.dll**：C# .NET Framework 4.8 COM 加载项（双位数注册，兼容 32/64 位 Office）

## 🗺️ 路线图

| 里程碑 | 内容 | 状态 |
|--------|------|------|
| M0 · 仓库骨架 | 资产迁移、构建链 | ✅ |
| M1 · 安全加固 + 托盘化 | token 鉴权 / CORS / agent.exe / 托盘 | ✅ |
| M2 · COM 薄壳 | Ribbon + WebView2 任务窗格 + 注册 | ✅ |
| M3 · 工具层迁移 | 24 工具 Word OM 化 + 34 用例回归 | ✅ |
| M4 · 统一安装器 | Inno Setup 双模式安装 / 零残留卸载 | ✅ |
| **批次 A · 体验包** | 自动标题 / 检查点回退 / 用量面板 | ✅ |
| 批次 B · 能力包 | 文档截图自检 / 调试 CLI / 联网搜索 | 🔜 |
| 批次 C · 生态包 | 订阅登录 / Skills / OOXML / VFS | 📋 |
| M5 · 验证矩阵 | 多版本 Office × 位数 × 系统覆盖 | 📋 |

## 🔐 安全模型

- 服务仅监听 `127.0.0.1`，所有 `/v1/*` 需 `X-Agent-Token`（启动时随机生成，文件 ACL 收紧至当前用户）
- CORS 固定白名单，绝不反射任意 Origin —— 恶意网页无法读取知识库 / 消耗额度
- 模型 API Key 仅存本机 `%LOCALAPPDATA%\OfficeAgent\`，面板中脱敏显示
- 卸载零残留（COM 注册 / 自启键 / 快捷方式全清），用户数据可选保留

## 🧑‍💻 开发指南

<details>
<summary>目录结构</summary>

```
apps/
├── word-addin/    # TS UI（聊天/知识库/检查点/用量），WebView2 加载
├── local-agent/   # Express 服务 → 打包为 agent.exe
├── com-host/      # C# COM 加载项（Ribbon + WebView2 桥 + Word OM 执行器）
└── tray/          # C# WinForms 托盘壳
installer/         # Inno Setup 安装器脚本
scripts/           # 打包 / 冒烟 / 工具回归 / COM 注册
plans/             # 设计文档与历史记录
```
</details>

<details>
<summary>测试与验证</summary>

- `npm run smoke`：服务健康 + token 鉴权 + CORS 冒烟
- `npm run test:tools`：真实 Word 实例驱动 24 个文档工具（34 用例，含空文档 / 中文 / 表格 / 流程图 / 大文档性能）
- 诊断日志：`%LOCALAPPDATA%\OfficeAgent\logs\com-host.log`（COM 生命周期 + 桥调用耗时）、`data\agent-trace.jsonl`（Agent 全链路）

</details>

## 🙏 致谢

- [office-agents](https://github.com/hewliyang/office-agents) — 文档截图自检、调试桥等特性的灵感来源
- [WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) · [Inno Setup](https://jrsoftware.org/isinfo.php) · [esbuild](https://esbuild.github.io/)

## 📄 License

开源协议待定（计划采用 MIT）。

---

<div align="center">
  <sub>Built with ❤️ for Word power users · Issues / PR 欢迎</sub>
</div>
