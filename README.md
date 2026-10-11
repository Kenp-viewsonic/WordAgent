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
npm run package           # 产出 dist/（agent.exe + 托盘 + public/ + com/）
npm run smoke             # 安全与健康冒烟（token 鉴权 / CORS / 启动）
```

## COM 加载项（WP2）

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register-com.ps1    # 注册（免管理员，自动识别 Office 位数）
powershell -ExecutionPolicy Bypass -File scripts\unregister-com.ps1  # 注销
```

注册后重启 Word，「开始」选项卡出现 **Word Agent → 打开助手** 按钮，点击打开任务窗格（WebView2 加载 UI）。
加载诊断日志：`%LOCALAPPDATA%\OfficeAgent\logs\com-host.log`。
构建机器需要 `apps\com-host\lib\Extensibility.dll`（已随源码内置，仅编译期使用；类型内嵌，运行时无 PIA 依赖）。

## 端到端验收清单（WP2）

1. 关闭所有 Word 窗口后重新打开，新建空白文档
2. 「开始」选项卡右侧出现 **Word Agent** 分组与 **打开助手** 按钮（无任何信任中心配置）
3. 点击按钮 → 右侧任务窗格打开：首次自动拉起托盘（右下角图标）→ 服务就绪 → 加载 UI
4. 发一条聊天消息（provider 配置在 `%LOCALAPPDATA%\OfficeAgent\data`，流式回复正常）
5. 知识库面板可上传/查看/删除文件
6. **预期不可用**：文档修改类工具（WP3 迁移中），调用会提示不在 Word 宿主/工具失败

排查：按钮缺失 → 重新运行 `register-com.ps1`，或 文件→选项→加载项→管理:禁用项目 中恢复；
面板异常 → 看 `%LOCALAPPDATA%\OfficeAgent\logs\com-host.log`。改代码重新构建前先关闭 Word（DLL 被占用）。

## 安全模型（WP1）

- `/v1/*` 全部要求 `X-Agent-Token` 请求头（`/health` 除外，供托盘探活）
- token 于服务启动时生成并存 `%LOCALAPPDATA%\OfficeAgent\agent.token`
- CORS 固定白名单（`AGENT_CORS_ORIGINS` 可扩展），不反射任意 Origin
