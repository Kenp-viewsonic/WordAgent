using System;
using System.Runtime.InteropServices;

namespace OfficeAgent.ComHost
{
    // Office COM interface identities declared locally (no Office PIA needed
    // for these two). GUIDs and signatures verified against the Office 15 PIA
    // (Microsoft.Office.Core). IDTExtensibility2 comes from the Extensibility
    // PIA instead — see the csproj reference (embedded interop types).

    /// <summary>Microsoft.Office.Core.IRibbonExtensibility (000C0396-0000-0000-C000-000000000046).</summary>
    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }

    /// <summary>Microsoft.Office.Core.ICustomTaskPaneConsumer (000C033E-0000-0000-C000-000000000046).</summary>
    [ComVisible(true)]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] object ctpFactoryInst);
    }
}
