using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeAgent.ComHost
{
    /// <summary>Embedded PNG → IPictureDisp for Ribbon loadImage callbacks.</summary>
    internal static class IconLoader
    {
        private const int PictypeBitmap = 1;
        private static readonly Guid IidIPictureDisp = new Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB");

        [StructLayout(LayoutKind.Sequential)]
        private struct PictDesc
        {
            public int cbSizeOfStruct;
            public int picType;
            public IntPtr hbitmap;
            public IntPtr hpalette;
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void OleCreatePictureIndirect(ref PictDesc pictDesc, ref Guid iid, bool fOwn, out object ppvObj);

        public static object GetPictureDisp(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string fullName = "OfficeAgent.ComHost.Resources." + resourceName;
            using (Stream stream = asm.GetManifestResourceStream(fullName))
            {
                if (stream == null) return null;
                using (var bitmap = new Bitmap(stream))
                {
                    var desc = new PictDesc
                    {
                        cbSizeOfStruct = Marshal.SizeOf(typeof(PictDesc)),
                        picType = PictypeBitmap,
                        hbitmap = bitmap.GetHbitmap(), // copy owned by OLE (fOwn below)
                        hpalette = IntPtr.Zero,
                    };
                    Guid iid = IidIPictureDisp;
                    OleCreatePictureIndirect(ref desc, ref iid, true, out object picture);
                    return picture;
                }
            }
        }
    }
}
