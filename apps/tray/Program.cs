using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OfficeAgent.Tray
{
    /// <summary>
    /// Tray-resident supervisor for the local agent (WP1 of
    /// plans/com-shell-migration-kickoff.md): single instance, auto-start,
    /// port-conflict handling, service restart, log access.
    /// </summary>
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            using (var mutex = new Mutex(true, "OfficeAgent.Tray.SingleInstance", out bool createdNew))
            {
                if (!createdNew)
                {
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayContext());

                GC.KeepAlive(mutex);
            }
        }
    }

    internal sealed class TrayContext : ApplicationContext
    {
        private const string ServiceUrl = "http://127.0.0.1:8787/";
        private const string HealthUrl = "http://127.0.0.1:8787/health";
        private const int ServicePort = 8787;
        private const string AutostartKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AutostartValueName = "OfficeAgent";

        private static readonly HttpClient Http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(2),
        };

        private readonly NotifyIcon tray;
        private readonly ToolStripMenuItem autostartItem;
        private readonly System.Windows.Forms.Timer supervisor;
        private Process agentProcess;
        private int failedChecks;
        private bool exiting;

        public TrayContext()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("打开面板(&O)", null, delegate { OpenPanel(); });
            menu.Items.Add("重启服务(&R)", null, delegate { RestartService(); });

            autostartItem = new ToolStripMenuItem("开机自启(&A)");
            autostartItem.CheckOnClick = true;
            autostartItem.Checked = IsAutostartEnabled();
            autostartItem.CheckedChanged += delegate { SetAutostart(autostartItem.Checked); };
            menu.Items.Add(autostartItem);

            menu.Items.Add("查看日志(&L)", null, delegate { OpenLogs(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("退出(&X)", null, delegate { ExitApp(); });

            tray = new NotifyIcon
            {
                Icon = LoadIcon(),
                Text = "Word Agent",
                ContextMenuStrip = menu,
                Visible = true,
            };
            tray.DoubleClick += delegate { OpenPanel(); };

            supervisor = new System.Windows.Forms.Timer { Interval = 5000 };
            supervisor.Tick += async delegate { await CheckServiceAsync(); };
            supervisor.Start();

            Task.Run((Action)StartService);
        }

        private static Icon LoadIcon()
        {
            try
            {
                return Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application;
            }
            catch
            {
                return SystemIcons.Application;
            }
        }

        // ── Service supervision ────────────────────────────────────────────

        private void StartService()
        {
            if (exiting)
            {
                return;
            }

            if (IsServiceAliveAsync().GetAwaiter().GetResult())
            {
                failedChecks = 0;
                return;
            }

            FreePortIfOccupied();

            var exe = Path.Combine(AppContext.BaseDirectory, "agent.exe");
            if (!File.Exists(exe))
            {
                ShowBalloon("找不到 agent.exe，请检查安装目录。", ToolTipIcon.Error);
                return;
            }

            try
            {
                var psi = new ProcessStartInfo(exe)
                {
                    WorkingDirectory = AppContext.BaseDirectory,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                agentProcess = Process.Start(psi);
                if (agentProcess != null)
                {
                    agentProcess.EnableRaisingEvents = true;
                    agentProcess.Exited += delegate
                    {
                        agentProcess = null;
                        if (!exiting)
                        {
                            ShowBalloon("本地服务已退出，正在自动重启…", ToolTipIcon.Warning);
                            Task.Run((Action)StartService);
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                ShowBalloon("启动本地服务失败：" + ex.Message, ToolTipIcon.Error);
            }
        }

        private async Task CheckServiceAsync()
        {
            if (exiting)
            {
                return;
            }

            if (await IsServiceAliveAsync())
            {
                failedChecks = 0;
                return;
            }

            failedChecks++;
            if (failedChecks <= 30)
            {
                _ = Task.Run((Action)StartService);
            }
        }

        private static async Task<bool> IsServiceAliveAsync()
        {
            try
            {
                var res = await Http.GetAsync(HealthUrl);
                return res.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Port-conflict handling ported from scripts/start.bat: if 8787 is
        /// held by a foreign process, kill it so the agent can bind.
        /// </summary>
        private void FreePortIfOccupied()
        {
            foreach (int pid in GetListeningPids(ServicePort))
            {
                if (pid == Process.GetCurrentProcess().Id)
                {
                    continue;
                }
                try
                {
                    using (var proc = Process.GetProcessById(pid))
                    {
                        proc.Kill();
                        ShowBalloon(
                            string.Format("端口 {0} 被 {1} 占用，已结束该进程。", ServicePort, proc.ProcessName),
                            ToolTipIcon.Warning);
                    }
                }
                catch
                {
                    // process already gone or access denied — ignore
                }
            }
        }

        private static IEnumerable<int> GetListeningPids(int port)
        {
            var pids = new List<int>();
            try
            {
                var psi = new ProcessStartInfo("netstat.exe", "-ano -p tcp")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                };
                using (var proc = Process.Start(psi))
                {
                    string line;
                    while ((line = proc.StandardOutput.ReadLine()) != null)
                    {
                        var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length < 5)
                        {
                            continue;
                        }
                        if (parts[1].EndsWith(":" + port) &&
                            parts[3].Equals("LISTENING", StringComparison.OrdinalIgnoreCase) &&
                            int.TryParse(parts[4], out int pid))
                        {
                            pids.Add(pid);
                        }
                    }
                    proc.WaitForExit();
                }
            }
            catch
            {
                // netstat unavailable — skip conflict handling
            }
            return pids;
        }

        // ── Menu actions ──────────────────────────────────────────────────

        private void OpenPanel()
        {
            try
            {
                Process.Start(ServiceUrl);
            }
            catch (Exception ex)
            {
                ShowBalloon("无法打开面板：" + ex.Message, ToolTipIcon.Error);
            }
        }

        private void RestartService()
        {
            Task.Run(() =>
            {
                try
                {
                    if (agentProcess != null && !agentProcess.HasExited)
                    {
                        agentProcess.Kill();
                        agentProcess = null;
                    }
                }
                catch
                {
                    // ignore
                }
                StartService();
            });
        }

        private void OpenLogs()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var trace = Path.Combine(localAppData, "OfficeAgent", "data", "agent-trace.jsonl");
            try
            {
                if (File.Exists(trace))
                {
                    Process.Start(trace);
                }
                else
                {
                    var dir = Path.GetDirectoryName(trace);
                    Directory.CreateDirectory(dir);
                    Process.Start("explorer.exe", dir);
                }
            }
            catch (Exception ex)
            {
                ShowBalloon("无法打开日志：" + ex.Message, ToolTipIcon.Error);
            }
        }

        private void ExitApp()
        {
            exiting = true;
            supervisor.Stop();
            try
            {
                if (agentProcess != null && !agentProcess.HasExited)
                {
                    agentProcess.Kill();
                }
            }
            catch
            {
                // ignore
            }
            tray.Visible = false;
            tray.Dispose();
            ExitThread();
        }

        // ── Autostart (HKCU Run) ──────────────────────────────────────────

        private static bool IsAutostartEnabled()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(AutostartKey, false))
                {
                    return key != null && key.GetValue(AutostartValueName) != null;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void SetAutostart(bool enabled)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(AutostartKey))
                {
                    if (enabled)
                    {
                        key.SetValue(AutostartValueName, "\"" + Application.ExecutablePath + "\"");
                    }
                    else
                    {
                        key.DeleteValue(AutostartValueName, false);
                    }
                }
            }
            catch
            {
                // registry unavailable — user can still start manually
            }
        }

        private void ShowBalloon(string text, ToolTipIcon icon)
        {
            try
            {
                tray.ShowBalloonTip(3000, "Word Agent", text, icon);
            }
            catch
            {
                // ignore
            }
        }
    }
}
