# Office Agent Desktop（预览构建）

## 运行

双击 `OfficeAgent.Tray.exe`：托盘常驻并拉起 `agent.exe`（http://127.0.0.1:8787）。
托盘菜单：打开面板 / 重启服务 / 开机自启 / 查看日志 / 退出。

前置：Windows 10/11 + .NET Framework 4.8（系统自带）。无需安装 Node.js。

## 说明

- 首次启动生成访问 token，UI 由本机服务托管并自动携带
- 数据存于 `%LOCALAPPDATA%\OfficeAgent\data`
- COM 加载项与统一安装器在后续里程碑（WP2/WP4）提供
