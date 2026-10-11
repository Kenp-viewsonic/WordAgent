; Office Agent Desktop — 统一安装器 (WP4, plans/com-shell-migration-kickoff.md §4)
;
; Build (after npm run package):
;   iscc installer\OfficeAgent.iss        (or: npm run build:installer)
;
; 安装模式（PrivilegesRequiredOverridesAllowed=dialog，启动时选择）：
;   所有用户: {autopf}\OfficeAgent + HKLM COM 注册（需管理员）
;   仅当前用户: %LOCALAPPDATA%\Programs\OfficeAgent + HKCU COM 注册（免管理员）
; 两种模式的注册布局与 scripts\register-com.ps1 一致（mscoree.dll + CodeBase，
; 32/64 位注册表视图双写，兼容 32/64 位 Office）。

#define MyAppName "Word Agent"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "OfficeAgent"
#define MyAppExeName "OfficeAgent.Tray.exe"
#define DistDir "..\dist"

; Pre-escaped CLSIDs ({{ = literal brace for Inno's constant parser)
#define ClsAddin "{{EC9522B6-2A5E-4E51-9B7D-55CA3C90A5D6}"
#define ClsPane "{{4342A9CF-0BA6-424E-AEF5-0134DDCD9DFA}"
#define IfaceLockback "{{000C0601-0000-0000-C000-000000000046}"

[Setup]
AppId={#ClsAddin}
AppName={#MyAppName}
AppVersion={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\OfficeAgent
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
OutputDir=..\dist-setup
OutputBaseFilename=OfficeAgent-Setup-{#MyAppVersion}
SetupIconFile=..\apps\tray\app.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; Flags: checkedonce

[Files]
Source: "{#DistDir}\agent.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#DistDir}\OfficeAgent.Tray.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#DistDir}\OfficeAgent.Tray.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#DistDir}\public\*"; DestDir: "{app}\public"; Flags: ignoreversion recursesubdirs
Source: "{#DistDir}\com\x64\*"; DestDir: "{app}\com\x64"; Flags: ignoreversion
Source: "{#DistDir}\com\x86\*"; DestDir: "{app}\com\x86"; Flags: ignoreversion
Source: "{#DistDir}\README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; ── COM class: Addin (64-bit view → x64 DLL) ────────────────────────────────
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}"; ValueType: string; ValueData: "OfficeAgent.ComHost.Addin"; Flags: uninsdeletekey
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OfficeAgent.ComHost.Addin"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "OfficeAgent.ComHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\com\x64\OfficeAgent.ComHost.dll"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsAddin}\Implemented Categories\{#IfaceLockback}"; ValueType: none

; ── COM class: Addin (32-bit view → x86 DLL, 32 位 Office) ──────────────────
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}"; ValueType: string; ValueData: "OfficeAgent.ComHost.Addin"; Flags: uninsdeletekey
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OfficeAgent.ComHost.Addin"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "OfficeAgent.ComHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\com\x86\OfficeAgent.ComHost.dll"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsAddin}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"

; ── COM class: Task pane control (双视图) ───────────────────────────────────
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}"; ValueType: string; ValueData: "OfficeAgent.ComHost.AgentTaskPane"; Flags: uninsdeletekey
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OfficeAgent.ComHost.AgentTaskPane"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "OfficeAgent.ComHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\com\x64\OfficeAgent.ComHost.dll"
Root: HKA64; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}"; ValueType: string; ValueData: "OfficeAgent.ComHost.AgentTaskPane"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OfficeAgent.ComHost.AgentTaskPane"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "OfficeAgent.ComHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\com\x86\OfficeAgent.ComHost.dll"
Root: HKA32; Subkey: "Software\Classes\CLSID\{#ClsPane}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"

; ── ProgIDs（双视图） ───────────────────────────────────────────────────────
Root: HKA64; Subkey: "Software\Classes\OfficeAgent.Addin"; ValueType: string; ValueData: "OfficeAgent.ComHost.Addin"; Flags: uninsdeletekey
Root: HKA64; Subkey: "Software\Classes\OfficeAgent.Addin\CLSID"; ValueType: string; ValueData: "{#ClsAddin}"
Root: HKA64; Subkey: "Software\Classes\OfficeAgent.TaskPane"; ValueType: string; ValueData: "OfficeAgent.ComHost.AgentTaskPane"; Flags: uninsdeletekey
Root: HKA64; Subkey: "Software\Classes\OfficeAgent.TaskPane\CLSID"; ValueType: string; ValueData: "{#ClsPane}"
Root: HKA32; Subkey: "Software\Classes\OfficeAgent.Addin"; ValueType: string; ValueData: "OfficeAgent.ComHost.Addin"
Root: HKA32; Subkey: "Software\Classes\OfficeAgent.Addin\CLSID"; ValueType: string; ValueData: "{#ClsAddin}"
Root: HKA32; Subkey: "Software\Classes\OfficeAgent.TaskPane"; ValueType: string; ValueData: "OfficeAgent.ComHost.AgentTaskPane"
Root: HKA32; Subkey: "Software\Classes\OfficeAgent.TaskPane\CLSID"; ValueType: string; ValueData: "{#ClsPane}"

; ── Word 加载项键 LoadBehavior=3（双视图，Office 按位数读取） ───────────────
Root: HKA64; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeAgent.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Word Agent"; Flags: uninsdeletekey
Root: HKA64; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeAgent.Addin"; ValueType: string; ValueName: "Description"; ValueData: "Word Agent 本地助手（COM 加载项）"
Root: HKA64; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeAgent.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKA32; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeAgent.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Word Agent"
Root: HKA32; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeAgent.Addin"; ValueType: string; ValueName: "Description"; ValueData: "Word Agent 本地助手（COM 加载项）"
Root: HKA32; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeAgent.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3

; ── KB948461: .NET Framework lockback bypass（双视图） ──────────────────────
Root: HKA64; Subkey: "Software\Classes\Interface\{#IfaceLockback}"; ValueType: string; ValueData: "Office .NET Framework Lockback Bypass Key"
Root: HKA32; Subkey: "Software\Classes\Interface\{#IfaceLockback}"; ValueType: string; ValueData: "Office .NET Framework Lockback Bypass Key"

; ── 开机自启（HKCU，托盘菜单可关） ─────────────────────────────────────────
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "OfficeAgent"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName} 托盘"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -WindowStyle Hidden -Command ""Get-Process | Where-Object {{ $_.Path -like '{app}\*' }} | Stop-Process -Force -ErrorAction SilentlyContinue"""; RunOnceId: "StopOfficeAgent"; Flags: runhidden

[Code]
function FindWindowA(lpClassName: AnsiString; lpWindowName: AnsiString): LongInt;
  external 'FindWindowA@user32.dll stdcall';

function WordIsRunning(): Boolean;
begin
  // Word's main window class is "OpusApp"
  Result := FindWindowA('OpusApp', '') <> 0;
end;

function WebView2Installed(): Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(HKLM,
    'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version)
    and (Version <> '');
  if not Result then
    Result := RegQueryStringValue(HKLM,
      'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version)
      and (Version <> '');
end;

function InitializeSetup(): Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;

  // Word holds the COM DLL while running — without this check, the silent
  // "close applications" prompt defaults to Abort and the install rolls back.
  if WordIsRunning() then
  begin
    if WizardSilent() then
      Result := False
    else
    begin
      MsgBox('检测到 Microsoft Word 正在运行。' + #13#10 + #13#10 +
        '安装需要更新被 Word 占用的文件，请先完全退出 Word（含后台进程）后重试。',
        mbError, MB_OK);
      Result := False;
    end;
    exit;
  end;

  if not WebView2Installed() then
  begin
    if MsgBox('Word Agent 任务窗格需要 Microsoft Edge WebView2 运行时' + #13#10 +
      '（Win10/11 + 新版 Office 通常已自带）。' + #13#10 + #13#10 +
      '是否现在下载并安装 WebView2？', mbConfirmation, MB_YESNO) = IDYES then
      ShellExec('open', 'https://go.microsoft.com/fwlink/p/?LinkId=2124703',
        '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  if WordIsRunning() then
  begin
    if UninstallSilent() then
      Result := False
    else
    begin
      MsgBox('检测到 Microsoft Word 正在运行。' + #13#10 + #13#10 +
        '卸载需要移除被 Word 占用的文件，请先完全退出 Word（含后台进程）后重试。',
        mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  // Silent uninstall keeps user data (no prompt possible) — /SUPPRESSMSGBOXES
  // does not suppress custom MsgBox calls.
  if (CurUninstallStep = usUninstall) and (not UninstallSilent()) then
  begin
    DataDir := ExpandConstant('{localappdata}\OfficeAgent');
    if DirExists(DataDir) then
    begin
      if MsgBox('是否同时删除用户数据（模型配置 / 知识库 / 会话日志）？' + #13#10 + DataDir,
        mbConfirmation, MB_YESNO) = IDYES then
        DelTree(DataDir, True, True, True);
    end;
  end;
end;
