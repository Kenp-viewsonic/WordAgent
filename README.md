# Office Agent Desktop

> 重构进行中：决策记录、现状清点与工作包分解见 [plans/com-shell-migration-kickoff.md](plans/com-shell-migration-kickoff.md)（唯一信息入口）。

Windows 桌面版 Word 本地 Agent：C# COM 薄壳加载项 + WebView2 任务窗格 + 托盘常驻本地服务（agent.exe）。

## 目录结构

| 路径 | 说明 |
|------|------|
| `apps/word-addin` | TS UI（聊天/知识库/预设），由 agent.exe 托管，WebView2 中加载 |
| `apps/local-agent` | 本地 Agent 服务（Express），打包为 `agent.exe` |
| `apps/com-host` | C# .NET Framework 4.8 COM 加载项（WP2/WP3） |
| `apps/tray` | C# WinForms 托盘壳（单实例/自启/服务守护） |
| `installer/` | Inno Setup 安装器（WP4） |
| `scripts/` | 打包与冒烟脚本 |
| `plans/` | 计划与历史文档 |

`officeAgent/` 仅为历史参考（旧 Web Add-in 线），不参与构建、不入库。

## 开发

```bash
npm install
npm run dev:agent     # 本地 Agent：http://127.0.0.1:8787
npm run dev:addin     # UI 开发服务器：http://localhost:3001（代理 /api → 8787）
```

先启动 `dev:agent`（生成 token 文件）再启动 `dev:addin`。数据与 token 存于 `%LOCALAPPDATA%\OfficeAgent\`。

## 构建 / 打包 / 冒烟

```bash
npm run build:prod        # TS + UI 构建
npm run build:agent-exe   # esbuild + Node SEA → build/agent.exe
npm run package           # 产出 dist/（agent.exe + OfficeAgent.Tray.exe + public/）
npm run smoke             # 安全与健康冒烟（token 鉴权 / CORS / 启动）
```

## 安全模型（WP1）

- `/v1/*` 全部要求 `X-Agent-Token` 请求头（`/health` 除外，供托盘探活）
- token 于服务启动时生成并存 `%LOCALAPPDATA%\OfficeAgent\agent.token`
- CORS 固定白名单（`AGENT_CORS_ORIGINS` 可扩展），不反射任意 Origin
