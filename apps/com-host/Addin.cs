using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Win32;

namespace OfficeAgent.ComHost
{
    /// <summary>
    /// COM thin shell loaded by Word (decision D1): Ribbon button + task pane
    /// host. Registered with regasm /codebase; the Word Addins key
    /// (LoadBehavior=3) is written by RegisterFunction below.
    /// See plans/com-shell-migration-kickoff.md §6.1/§6.2.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    // AutoDual: Ribbon onAction/loadImage/onLoad callbacks are dispatched by
    // name through IDispatch — with ClassInterface(None) the class methods are
    // absent from the name table and Word silently drops the click.
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Addin : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        public const string ClassId = "EC9522B6-2A5E-4E51-9B7D-55CA3C90A5D6";
        public const string ProgId = "OfficeAgent.Addin";

        private const string AddinsKeyPath = @"Software\Microsoft\Office\Word\Addins\";
        private const string LockbackInterfacePath = @"Interface\{000C0601-0000-0000-C000-000000000046}";

        private static object taskPane; // Microsoft.Office.Core.CustomTaskPane (late-bound)
        private static AgentTaskPane paneControl;

        private const string RibbonXml = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoaded"" loadImage=""LoadImage"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""oaGroup"" label=""Word Agent"">
          <button id=""oaOpen"" label=""打开助手"" size=""large"" image=""oaIcon32"" onAction=""OnOpenPanel"" screentip=""Word Agent"" supertip=""打开 Word Agent 助手任务窗格""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public Addin()
        {
            Trace("Addin ctor");
        }

        static Addin()
        {
            // Surface marshaling/interop failures that Word swallows silently.
            // Only interop-shaped exceptions are logged — first-chance noise
            // from background HTTP probes would drown the signal otherwise.
            AppDomain.CurrentDomain.FirstChanceException += (sender, e) =>
            {
                string name = e.Exception.GetType().Name;
                if (name.Contains("COM") || name.Contains("Marshal") || name.Contains("SEH") ||
                    name.Contains("TypeLoad") || name.Contains("FileNotFound") || name.Contains("BadImage"))
                {
                    Trace("FirstChance: " + e.Exception.GetType().FullName + ": " + e.Exception.Message);
                }
            };
        }

        // ── IDTExtensibility2 ─────────────────────────────────────────────────

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            // The Word Application object is late-bound in WP3 (Word OM tool
            // executor); WP2 only needs the connection lifecycle.
            Trace("OnConnection connectMode=" + connectMode);
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Trace("OnDisconnection removeMode=" + removeMode);
            try
            {
                if (paneControl != null)
                {
                    paneControl.Dispose();
                    paneControl = null;
                }
            }
            catch
            {
                // pane may already be gone
            }
            taskPane = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ── IRibbonExtensibility ──────────────────────────────────────────────

        public string GetCustomUI(string ribbonID)
        {
            Trace("GetCustomUI ribbonID=" + ribbonID);
            return RibbonXml;
        }

        public void OnRibbonLoaded(object ribbon)
        {
            Trace("OnRibbonLoaded");
        }

        public object LoadImage(string imageId)
        {
            Trace("LoadImage " + imageId);
            return IconLoader.GetPictureDisp("icon-32.png");
        }

        public void OnOpenPanel(object control)
        {
            Trace("OnOpenPanel taskPane=" + (taskPane != null));
            try
            {
                if (taskPane == null)
                {
                    MessageBox.Show(
                        "任务窗格尚未就绪（宿主未提供 CTP 工厂），请重启 Word 后重试。",
                        "Word Agent",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                dynamic ctp = taskPane;
                bool show = !ctp.Visible;
                Trace("OnOpenPanel toggling visible=" + show);
                if (show && paneControl != null)
                {
                    paneControl.EnsureStarted();
                }
                ctp.Visible = show;
            }
            catch (Exception ex)
            {
                Trace("OnOpenPanel failed: " + ex.GetType().Name + ": " + ex.Message);
                MessageBox.Show(
                    "打开任务窗格失败：" + ex.Message,
                    "Word Agent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ── ICustomTaskPaneConsumer ───────────────────────────────────────────

        public void CTPFactoryAvailable(object ctpFactoryInst)
        {
            try
            {
                dynamic factory = ctpFactoryInst;
                taskPane = factory.CreateCTP(AgentTaskPane.ProgId, "Word Agent", Type.Missing);
                dynamic ctp = taskPane;
                ctp.DockPosition = 2; // msoCTPDockPositionRight
                ctp.Width = 420;
                ctp.Visible = false; // the Ribbon button toggles it

                object content = ctp.ContentControl;
                paneControl = content as AgentTaskPane ?? AgentTaskPane.LastInstance;
                Trace("CTP created");
            }
            catch (Exception ex)
            {
                Trace("CTP failed: " + ex.Message);
                MessageBox.Show(
                    "创建任务窗格失败：" + ex.Message,
                    "Word Agent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ── diagnostics ───────────────────────────────────────────────────────

        internal static void Trace(string message)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAgent",
                    "logs");
                Directory.CreateDirectory(dir);
                File.AppendAllText(
                    Path.Combine(dir, "com-host.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message + Environment.NewLine);
            }
            catch
            {
                // tracing must never break the add-in
            }
        }

        // ── COM registration (executed by regasm) ─────────────────────────────

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            if (type != typeof(Addin)) return;

            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(AddinsKeyPath + ProgId))
            {
                key.SetValue("FriendlyName", "Word Agent");
                key.SetValue("Description", "Word Agent 本地助手（COM 加载项）");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }

            // KB948461: .NET Framework lockback bypass — without this key newer
            // Office builds may refuse to load .NET 4.x COM add-ins.
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(LockbackInterfacePath))
            {
                if (key.GetValue("") == null)
                {
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                }
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            if (type != typeof(Addin)) return;
            Registry.CurrentUser.DeleteSubKeyTree(AddinsKeyPath + ProgId, false);
        }
    }
}
