using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office    = Microsoft.Office.Core;

namespace PPTToolbox
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _customTaskPane;
        internal RibbonPPT     Ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var ctrl            = new TaskPaneControl();
            _customTaskPane     = CustomTaskPanes.Add(ctrl, BrandingConfig.ToolName);
            _customTaskPane.Width    = 300;
            _customTaskPane.Visible  = false;
            _customTaskPane.VisibleChanged += (s, ev) => Ribbon?.RefreshTogglePane();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        // Called by ribbon toggle button
        public void SetPaneVisible(bool visible)
        {
            if (_customTaskPane != null)
                _customTaskPane.Visible = visible;
        }

        public bool IsPaneVisible =>
            _customTaskPane != null && _customTaskPane.Visible;

        // Called by ribbon "Show Panel" buttons
        public void ShowPane()
        {
            if (_customTaskPane != null)
                _customTaskPane.Visible = true;
        }

        // Application field is declared in ThisAddIn.Designer.cs
        internal PowerPoint.Application PPTApp => this.Application;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new RibbonPPT();
            return Ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup  += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
