using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeAgent.ComHost
{
    /// <summary>
    /// Task pane content: WebView2 hosting the existing TS UI from the local
    /// agent (http://127.0.0.1:8787). Instantiated by the Office CTP factory
    /// through its ProgId, so it must stay COM-visible with a parameterless
    /// constructor. WebView2 profile data lives under
    /// %LOCALAPPDATA%\OfficeAgent\WebView2 (never in Program Files).
    /// </summary>
    [ComVisible(true)]
    [Guid(PaneClassId)]
    [ProgId(ProgId)]
    // AutoDual: the Office CTP factory QIs the control for IDispatch —
    // ClassInterface(None) left it without one ("Cannot create the specified
    // ActiveX control").
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class AgentTaskPane : UserControl
    {
        public const string PaneClassId = "4342A9CF-0BA6-424E-AEF5-0134DDCD9DFA";
        public const string ProgId = "OfficeAgent.TaskPane";

        internal static AgentTaskPane LastInstance;

        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
        private const string ServiceUrl = "http://127.0.0.1:8787/";
        private const string HealthUrl = "http://127.0.0.1:8787/health";

        private readonly WebView2 webView;
        private readonly Label message;
        private bool started;

        public AgentTaskPane()
        {
            LastInstance = this;
            Dock = DockStyle.Fill;
            BackColor = Color.White;

            message = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Padding = new Padding(16),
                Text = "正在启动本地服务…",
            };
            webView = new WebView2 { Dock = DockStyle.Fill, Visible = false };
            Controls.Add(webView);
            Controls.Add(message);
        }

        public void EnsureStarted()
        {
            if (started) return;
            started = true;
            Addin.Trace("Pane.EnsureStarted");
            _ = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                if (!await IsServiceUpAsync())
                {
                    Addin.Trace("Pane: service not up, starting tray");
                    StartTray();
                    for (int i = 0; i < 15; i++)
                    {
                        await Task.Delay(1000);
                        if (await IsServiceUpAsync()) break;
                    }
                }

                if (!await IsServiceUpAsync())
                {
                    Addin.Trace("Pane: service still down after wait");
                    message.Text = "本地服务未运行。\r\n请从开始菜单启动 Word Agent（托盘程序）后重试。";
                    return;
                }

                string profileDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAgent",
                    "WebView2");
                CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(null, profileDir);
                await webView.EnsureCoreWebView2Async(env);
                webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                webView.CoreWebView2.WebMessageReceived += OnWebMessage;
                webView.CoreWebView2.Navigate(ServiceUrl);
                webView.Visible = true;
                message.Visible = false;
                Addin.Trace("Pane WebView2 ready, navigated " + ServiceUrl);
            }
            catch (Exception ex)
            {
                Addin.Trace("Pane init failed: " + ex.Message);
                message.Text = "任务窗格初始化失败：" + ex.Message;
            }
        }

        // ── WP3 bridge: UI ⇄ Word OM tool execution ──────────────────────────

        private void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                JObject msg = JObject.Parse(e.WebMessageAsJson);
                if ((string)msg["type"] != "tool-request") return;

                int id = (int)msg["id"];
                string method = (string)msg["method"] ?? "";
                JObject payload = (JObject)msg["payload"] ?? new JObject();
                JObject action = payload["action"] as JObject;
                string actionName = action != null ? (string)action["action"] : "";
                Addin.Trace("Bridge call: " + method + (actionName.Length > 0 ? " (" + actionName + ")" : ""));
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                WordToolExecutor.ToolResult result = WordToolExecutor.Execute(method, payload);
                stopwatch.Stop();

                var response = new JObject
                {
                    ["type"] = "tool-response",
                    ["id"] = id,
                    ["success"] = result.Success,
                };
                if (result.Success)
                {
                    response["result"] = result.Result ?? "";
                    if (result.Data != null) response["data"] = JToken.FromObject(result.Data);
                }
                else
                {
                    response["error"] = result.Error ?? "文档工具执行失败";
                }
                webView.CoreWebView2.PostWebMessageAsJson(response.ToString(Formatting.None));
                Addin.Trace("Bridge done: " + method + " success=" + result.Success + " " + stopwatch.ElapsedMilliseconds + "ms");
            }
            catch (Exception ex)
            {
                Addin.Trace("Bridge handler failed: " + ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static async Task<bool> IsServiceUpAsync()
        {
            try
            {
                HttpResponseMessage res = await Http.GetAsync(HealthUrl);
                return res.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Service-not-running fallback (WP2 acceptance): try to bring up the
        /// tray program, which supervises agent.exe.
        /// </summary>
        private static void StartTray()
        {
            string dir = Path.GetDirectoryName(typeof(AgentTaskPane).Assembly.Location);
            string found = FindTray(dir);
            if (found != null)
            {
                Addin.Trace("Pane: starting tray " + found);
                Process.Start(new ProcessStartInfo(found)
                {
                    WorkingDirectory = Path.GetDirectoryName(found),
                    UseShellExecute = true,
                });
            }
            else
            {
                Addin.Trace("Pane: OfficeAgent.Tray.exe not found near " + dir);
            }
        }

        private static string FindTray(string startDir)
        {
            string fromEnv = Environment.GetEnvironmentVariable("OFFICE_AGENT_TRAY");
            if (!string.IsNullOrEmpty(fromEnv) && File.Exists(fromEnv)) return fromEnv;

            DirectoryInfo dir = new DirectoryInfo(startDir);
            for (int depth = 0; depth < 6 && dir != null; depth++, dir = dir.Parent)
            {
                string[] candidates =
                {
                    // installed layout: next to the COM DLL (or its parent)
                    Path.Combine(dir.FullName, "OfficeAgent.Tray.exe"),
                    // package output: <repo>\dist
                    Path.Combine(dir.FullName, "dist", "OfficeAgent.Tray.exe"),
                    // dev build output
                    Path.Combine(dir.FullName, "apps", "tray", "bin", "Release", "net48", "OfficeAgent.Tray.exe"),
                };
                foreach (string candidate in candidates)
                {
                    if (File.Exists(candidate)) return candidate;
                }
            }
            return null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && webView != null)
            {
                webView.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
