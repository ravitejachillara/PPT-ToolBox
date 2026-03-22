using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace PPTToolbox
{
    [ComVisible(true)]
    public class RibbonPPT : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        // ── IRibbonExtensibility ─────────────────────────────────────────────────
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPTToolbox.RibbonPPT.xml");
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // Causes the toggle button to re-query its pressed state
        public void RefreshTogglePane()
        {
            _ribbon?.InvalidateControl("tbnShowPane");
        }

        // ── Toggle pane button ───────────────────────────────────────────────────
        public void TogglePane_Click(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.SetPaneVisible(pressed);
        }

        public bool TogglePane_GetPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.IsPaneVisible;
        }

        // ── Custom ribbon icon ───────────────────────────────────────────────────
        public stdole.IPictureDisp GetRibbonIcon(Office.IRibbonControl control)
        {
            try
            {
                string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string iconPath = Path.Combine(folder, "Resources", "ribbon_icon.png");
                if (File.Exists(iconPath))
                {
                    using (var bmp = new Bitmap(iconPath))
                        return AxHostImageConverter.GetIPictureDisp(bmp);
                }
            }
            catch { }
            return null;
        }

        // ── Helpers ──────────────────────────────────────────────────────────────
        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return string.Empty;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        // Converts a Bitmap to IPictureDisp for ribbon icons
        private sealed class AxHostImageConverter : System.Windows.Forms.AxHost
        {
            private AxHostImageConverter() : base(string.Empty) { }

            public static stdole.IPictureDisp GetIPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
    }
}
