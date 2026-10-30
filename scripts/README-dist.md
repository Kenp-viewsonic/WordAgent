# Office Agent Desktop（桌面版）

## 运行

双击 `OfficeAgent-Setup-0.2.0.exe` 安装（可选"所有用户/仅当前用户"），
安装完成后打开 Word → 「开始」选项卡 → **Word Agent → 打开助手**。
托盘常驻：右下角图标菜单 = 打开面板 / 重启服务 / 开机自启 / 查看日志 / 退出。

前置：Windows 10/11 + .NET Framework 4.8（系统自带）+ WebView2 运行时（安装器会检测并引导）。
无需安装 Node.js。**安装/卸载前请完全退出 Word。**

## 说明

- 升级：直接运行新版本安装包覆盖安装即可，无需先卸载；配置、知识库与会话都会保留
- 数据存于 `%LOCALAPPDATA%\OfficeAgent\data`（卸载时可选择保留）
- 模型配置通过面板右上角设置（供应商 API Key 等）
- 卸载：开始菜单 → 卸载 Word Agent，或设置的应用列表
