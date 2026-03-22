using System.Drawing;
using System.Windows.Forms;

namespace PPTToolbox
{
    partial class AboutDialog
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        private Button btnClose;
        private Label lblTitle, lblVersion, lblWatermark, lblDesc;

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            Text            = "About " + BrandingConfig.ToolName;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition   = FormStartPosition.CenterScreen;
            Size            = new Size(340, 240);
            BackColor       = BrandingConfig.PanelBg;
            MaximizeBox     = false;
            MinimizeBox     = false;

            lblTitle = new Label
            {
                Text      = BrandingConfig.ToolName,
                Font      = new Font("Segoe UI", 16f, FontStyle.Bold),
                ForeColor = BrandingConfig.Accent,
                Location  = new Point(20, 20),
                AutoSize  = true,
            };

            lblVersion = new Label
            {
                Text      = "Version " + BrandingConfig.Version,
                Font      = new Font("Segoe UI", 9f),
                ForeColor = BrandingConfig.PanelFg,
                Location  = new Point(22, 60),
                AutoSize  = true,
            };

            lblDesc = new Label
            {
                Text      = "Professional PowerPoint formatting add-in for Windows.\n" +
                             "Designed for Office 2016 and later.",
                Font      = new Font("Segoe UI", 8.5f),
                ForeColor = BrandingConfig.PanelFg,
                Location  = new Point(22, 88),
                Size      = new Size(290, 48),
            };

            lblWatermark = new Label
            {
                Text      = BrandingConfig.Watermark,
                Font      = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = Color.FromArgb(160, 160, 190),
                Location  = new Point(22, 145),
                AutoSize  = true,
            };

            btnClose = new Button
            {
                Text      = "Close",
                Size      = new Size(80, 28),
                Location  = new Point(230, 172),
                FlatStyle = FlatStyle.Flat,
                BackColor = BrandingConfig.Accent,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 9f),
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += btnClose_Click;

            Controls.AddRange(new Control[] { lblTitle, lblVersion, lblDesc, lblWatermark, btnClose });
        }
    }
}
