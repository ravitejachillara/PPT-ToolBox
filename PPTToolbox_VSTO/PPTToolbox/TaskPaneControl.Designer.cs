using System.Drawing;
using System.Windows.Forms;

namespace PPTToolbox
{
    partial class TaskPaneControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        // ── All declared controls ────────────────────────────────────────────────
        // Tab 1 – Arrange & Size
        private Button btnAlignLeft, btnAlignCenterH, btnAlignRight;
        private Button btnAlignTop, btnAlignMiddleV, btnAlignBottom;
        private Button btnDistributeH, btnDistributeV;
        private Button btnBringForward, btnSendBackward, btnBringToFront, btnSendToBack;
        private Button btnMatchWidth, btnMatchHeight, btnMatchBoth;
        private TextBox txtWidth, txtHeight, txtX, txtY;
        private Button btnApplySize, btnApplyPos, btnReadGeometry;

        // Tab 2 – Font
        private ComboBox cmbFont, cmbFontSize;
        private Button btnApplyFont, btnBold, btnItalic, btnUnderline, btnFontColor;
        private FlowLayoutPanel pnlFontSwatches;

        // Tab 3 – Paragraph
        private Button btnParaLeft, btnParaCenter, btnParaRight, btnParaJustify;
        private TextBox txtLineSpacing, txtSpaceBefore, txtSpaceAfter;
        private Button btnApplySpacing;

        // Tab 4 – Fill & Outline
        private Button btnFillColor, btnNoFill;
        private FlowLayoutPanel pnlFillSwatches, pnlOutlineSwatches;
        private Button btnEditSwatches;
        private TrackBar trkTransparency;
        private Label lblTransparencyValue;
        private Button btnOutlineColor, btnApplyOutline, btnNoOutline;
        private TextBox txtOutlineWidth;

        // Tab 5 – Shadow & Quick
        private Button btnShadowSoft, btnShadowHard, btnShadowBottom, btnShadowPerspective, btnShadowRemove;
        private Button btnDuplicate, btnSave, btnAbout;

        // Shell
        private TabControl tabMain;
        private TabPage tpgArrange, tpgFont, tpgParagraph, tpgFill, tpgShadow;
        private Label lblHeader, lblWatermark;
        private Panel pnlHeader, pnlFooter;
        private PictureBox pbLogo;

        // ── InitializeComponent ──────────────────────────────────────────────────
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            // ── Colours ─────────────────────────────────────────────────────────
            Color bg      = BrandingConfig.PanelBg;
            Color fg      = BrandingConfig.PanelFg;
            Color accent  = BrandingConfig.Accent;
            Color secBg   = BrandingConfig.SectionBg;
            Color inBg    = BrandingConfig.InputBg;
            Color border  = BrandingConfig.BorderColor;
            Font  fntMain = new Font("Segoe UI", 8.5f);
            Font  fntBold = new Font("Segoe UI", 8.5f, FontStyle.Bold);

            // ── Helper lambdas ───────────────────────────────────────────────────
            Button MakeBtn(string text, int w = 60, int h = 24)
            {
                var b = new Button
                {
                    Text = text, Width = w, Height = h,
                    FlatStyle = FlatStyle.Flat, Font = fntMain,
                    BackColor = secBg, ForeColor = fg,
                    Margin = new Padding(2),
                    UseVisualStyleBackColor = false,
                };
                b.FlatAppearance.BorderColor = border;
                return b;
            }

            TextBox MakeTxt(int w = 55, string placeholder = "")
            {
                return new TextBox
                {
                    Width = w, Height = 22, Font = fntMain,
                    BackColor = inBg, ForeColor = fg,
                    BorderStyle = BorderStyle.FixedSingle,
                    Text = placeholder, Margin = new Padding(2),
                };
            }

            Label MakeLbl(string text, bool header = false)
            {
                return new Label
                {
                    Text = text, AutoSize = true, Font = header ? fntBold : fntMain,
                    ForeColor = header ? accent : fg, Margin = new Padding(2, 6, 2, 2),
                };
            }

            FlowLayoutPanel MakeSwatchPanel()
            {
                return new FlowLayoutPanel
                {
                    AutoSize = true, FlowDirection = FlowDirection.LeftToRight,
                    Margin = new Padding(2),
                };
            }

            // ════════════════════════════════════════════════════════════════════
            //  HEADER
            // ════════════════════════════════════════════════════════════════════
            pnlHeader = new Panel { Dock = DockStyle.Top, Height = 36, BackColor = BrandingConfig.Primary };
            lblHeader = new Label
            {
                Text = "  \u25A0 " + BrandingConfig.ToolName,
                Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = Color.White, TextAlign = ContentAlignment.MiddleLeft,
            };
            pnlHeader.Controls.Add(lblHeader);

            // ════════════════════════════════════════════════════════════════════
            //  FOOTER (logo + watermark)
            // ════════════════════════════════════════════════════════════════════
            pnlFooter = new Panel { Dock = DockStyle.Bottom, Height = 28, BackColor = BrandingConfig.Primary };

            // Logo: left-docked, initially hidden (width=0); LoadLogo() sets width if image found
            pbLogo = new PictureBox
            {
                Dock = DockStyle.Left,
                Width = 0,
                Height = 20,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = BrandingConfig.Primary,
                Margin = new Padding(4, 4, 2, 4),
                Visible = false,
            };

            lblWatermark = new Label
            {
                Text = BrandingConfig.Watermark,
                Dock = DockStyle.Fill, Font = new Font("Segoe UI", 7f, FontStyle.Italic),
                ForeColor = Color.FromArgb(160, 160, 190),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0),
            };

            // Add logo first (Left), then label (Fill fills remaining space)
            pnlFooter.Controls.Add(lblWatermark);
            pnlFooter.Controls.Add(pbLogo);

            // ════════════════════════════════════════════════════════════════════
            //  TAB CONTROL
            // ════════════════════════════════════════════════════════════════════
            tabMain = new TabControl
            {
                Dock = DockStyle.Fill, Font = fntMain,
                Padding = new Point(6, 4),
            };

            tpgArrange   = new TabPage("Arrange");
            tpgFont      = new TabPage("Font");
            tpgParagraph = new TabPage("Para");
            tpgFill      = new TabPage("Fill");
            tpgShadow    = new TabPage("Shadow");

            foreach (var tp in new[] { tpgArrange, tpgFont, tpgParagraph, tpgFill, tpgShadow })
            {
                tp.BackColor = bg;
                tp.ForeColor = fg;
                tabMain.TabPages.Add(tp);
            }

            // ════════════════════════════════════════════════════════════════════
            //  TAB 1: Arrange & Size
            // ════════════════════════════════════════════════════════════════════
            var flp1 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false, BackColor = bg,
                Padding = new Padding(4),
            };

            flp1.Controls.Add(MakeLbl("— Align to Slide —", true));

            var rowAlignH = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnAlignLeft    = MakeBtn("◄ Left",  62);
            btnAlignCenterH = MakeBtn("◈ Ctr H", 62);
            btnAlignRight   = MakeBtn("Right ►", 62);
            btnAlignLeft.Click    += btnAlignLeft_Click;
            btnAlignCenterH.Click += btnAlignCenterH_Click;
            btnAlignRight.Click   += btnAlignRight_Click;
            rowAlignH.Controls.AddRange(new Control[] { btnAlignLeft, btnAlignCenterH, btnAlignRight });

            var rowAlignV = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnAlignTop     = MakeBtn("▲ Top",   62);
            btnAlignMiddleV = MakeBtn("◈ Mid V", 62);
            btnAlignBottom  = MakeBtn("Bot ▼",   62);
            btnAlignTop.Click     += btnAlignTop_Click;
            btnAlignMiddleV.Click += btnAlignMiddleV_Click;
            btnAlignBottom.Click  += btnAlignBottom_Click;
            rowAlignV.Controls.AddRange(new Control[] { btnAlignTop, btnAlignMiddleV, btnAlignBottom });

            flp1.Controls.Add(rowAlignH);
            flp1.Controls.Add(rowAlignV);

            flp1.Controls.Add(MakeLbl("— Distribute —", true));
            var rowDist = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnDistributeH = MakeBtn("⇔ Dist H", 90);
            btnDistributeV = MakeBtn("⇕ Dist V", 90);
            btnDistributeH.Click += btnDistributeH_Click;
            btnDistributeV.Click += btnDistributeV_Click;
            rowDist.Controls.AddRange(new Control[] { btnDistributeH, btnDistributeV });
            flp1.Controls.Add(rowDist);

            flp1.Controls.Add(MakeLbl("— Z-Order —", true));
            var rowZ = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnBringForward = MakeBtn("Fwd", 46);
            btnSendBackward = MakeBtn("Back", 46);
            btnBringToFront = MakeBtn("Front", 50);
            btnSendToBack   = MakeBtn("To Back", 58);
            btnBringForward.Click += btnBringForward_Click;
            btnSendBackward.Click += btnSendBackward_Click;
            btnBringToFront.Click += btnBringToFront_Click;
            btnSendToBack.Click   += btnSendToBack_Click;
            rowZ.Controls.AddRange(new Control[] { btnBringForward, btnSendBackward, btnBringToFront, btnSendToBack });
            flp1.Controls.Add(rowZ);

            flp1.Controls.Add(MakeLbl("— Match Size —", true));
            var rowMatch = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnMatchWidth  = MakeBtn("Match W",  70);
            btnMatchHeight = MakeBtn("Match H",  70);
            btnMatchBoth   = MakeBtn("Match Both", 80);
            btnMatchWidth.Click  += btnMatchWidth_Click;
            btnMatchHeight.Click += btnMatchHeight_Click;
            btnMatchBoth.Click   += btnMatchBoth_Click;
            rowMatch.Controls.AddRange(new Control[] { btnMatchWidth, btnMatchHeight, btnMatchBoth });
            flp1.Controls.Add(rowMatch);

            flp1.Controls.Add(MakeLbl("— Exact Size (cm) —", true));
            var rowWH = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            rowWH.Controls.Add(MakeLbl("W:"));
            txtWidth  = MakeTxt(50, "0.00");
            rowWH.Controls.Add(txtWidth);
            rowWH.Controls.Add(MakeLbl("H:"));
            txtHeight = MakeTxt(50, "0.00");
            rowWH.Controls.Add(txtHeight);
            btnApplySize = MakeBtn("Apply", 50);
            btnApplySize.Click += btnApplySize_Click;
            rowWH.Controls.Add(btnApplySize);
            flp1.Controls.Add(rowWH);

            flp1.Controls.Add(MakeLbl("— Position (cm) —", true));
            var rowXY = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            rowXY.Controls.Add(MakeLbl("X:"));
            txtX = MakeTxt(50, "0.00");
            rowXY.Controls.Add(txtX);
            rowXY.Controls.Add(MakeLbl("Y:"));
            txtY = MakeTxt(50, "0.00");
            rowXY.Controls.Add(txtY);
            btnApplyPos = MakeBtn("Apply", 50);
            btnApplyPos.Click += btnApplyPos_Click;
            rowXY.Controls.Add(btnApplyPos);
            flp1.Controls.Add(rowXY);

            btnReadGeometry = MakeBtn("↓ Read from selection", 180, 26);
            btnReadGeometry.Click += btnReadGeometry_Click;
            flp1.Controls.Add(btnReadGeometry);

            tpgArrange.Controls.Add(flp1);

            // ════════════════════════════════════════════════════════════════════
            //  TAB 2: Font
            // ════════════════════════════════════════════════════════════════════
            var flp2 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false, BackColor = bg,
                Padding = new Padding(4),
            };

            flp2.Controls.Add(MakeLbl("— Font Family —", true));
            cmbFont = new ComboBox
            {
                Width = 175, Height = 24, Font = fntMain,
                BackColor = inBg, ForeColor = fg, DropDownStyle = ComboBoxStyle.DropDown,
                Margin = new Padding(2),
            };
            flp2.Controls.Add(cmbFont);

            flp2.Controls.Add(MakeLbl("— Size (pt) & Style —", true));
            var rowFontCtrl = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            cmbFontSize = new ComboBox
            {
                Width = 55, Height = 24, Font = fntMain,
                BackColor = inBg, ForeColor = fg, DropDownStyle = ComboBoxStyle.DropDown,
                Margin = new Padding(2),
            };
            btnBold      = MakeBtn("B",  28); btnBold.Font = fntBold;
            btnItalic    = MakeBtn("I",  28); btnItalic.Font = new Font("Segoe UI", 8.5f, FontStyle.Italic);
            btnUnderline = MakeBtn("U",  28); btnUnderline.Font = new Font("Segoe UI", 8.5f, FontStyle.Underline);
            btnBold.Click      += btnBold_Click;
            btnItalic.Click    += btnItalic_Click;
            btnUnderline.Click += btnUnderline_Click;
            rowFontCtrl.Controls.AddRange(new Control[] { cmbFontSize, btnBold, btnItalic, btnUnderline });
            flp2.Controls.Add(rowFontCtrl);

            flp2.Controls.Add(MakeLbl("— Font Colour —", true));
            var rowFontColor = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnFontColor = MakeBtn("■ Colour", 80, 26);
            btnFontColor.BackColor = Color.Black;
            btnFontColor.ForeColor = Color.White;
            btnFontColor.Click += btnFontColor_Click;
            btnApplyFont = MakeBtn("Apply All", 80, 26);
            btnApplyFont.BackColor = accent;
            btnApplyFont.ForeColor = Color.White;
            btnApplyFont.Click += btnApplyFont_Click;
            rowFontColor.Controls.AddRange(new Control[] { btnFontColor, btnApplyFont });
            flp2.Controls.Add(rowFontColor);

            flp2.Controls.Add(MakeLbl("— Colour Swatches —", true));
            pnlFontSwatches = MakeSwatchPanel();
            flp2.Controls.Add(pnlFontSwatches);

            tpgFont.Controls.Add(flp2);

            // ════════════════════════════════════════════════════════════════════
            //  TAB 3: Paragraph
            // ════════════════════════════════════════════════════════════════════
            var flp3 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false, BackColor = bg,
                Padding = new Padding(4),
            };

            flp3.Controls.Add(MakeLbl("— Text Alignment —", true));
            var rowParaAlign = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnParaLeft    = MakeBtn("≡L", 44);
            btnParaCenter  = MakeBtn("≡C", 44);
            btnParaRight   = MakeBtn("≡R", 44);
            btnParaJustify = MakeBtn("≡J", 44);
            btnParaLeft.Click    += btnParaLeft_Click;
            btnParaCenter.Click  += btnParaCenter_Click;
            btnParaRight.Click   += btnParaRight_Click;
            btnParaJustify.Click += btnParaJustify_Click;
            rowParaAlign.Controls.AddRange(new Control[] { btnParaLeft, btnParaCenter, btnParaRight, btnParaJustify });
            flp3.Controls.Add(rowParaAlign);

            flp3.Controls.Add(MakeLbl("— Line Spacing (x) —", true));
            var rowLS = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            txtLineSpacing = MakeTxt(55, "1.0");
            rowLS.Controls.Add(MakeLbl("Lines:"));
            rowLS.Controls.Add(txtLineSpacing);
            flp3.Controls.Add(rowLS);

            flp3.Controls.Add(MakeLbl("— Space Before / After (pt) —", true));
            var rowSBA = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            rowSBA.Controls.Add(MakeLbl("Bef:"));
            txtSpaceBefore = MakeTxt(50, "0");
            rowSBA.Controls.Add(txtSpaceBefore);
            rowSBA.Controls.Add(MakeLbl("Aft:"));
            txtSpaceAfter = MakeTxt(50, "0");
            rowSBA.Controls.Add(txtSpaceAfter);
            flp3.Controls.Add(rowSBA);

            btnApplySpacing = MakeBtn("Apply Spacing", 140, 26);
            btnApplySpacing.BackColor = accent;
            btnApplySpacing.ForeColor = Color.White;
            btnApplySpacing.Click += btnApplySpacing_Click;
            flp3.Controls.Add(btnApplySpacing);

            tpgParagraph.Controls.Add(flp3);

            // ════════════════════════════════════════════════════════════════════
            //  TAB 4: Fill & Outline
            // ════════════════════════════════════════════════════════════════════
            var flp4 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false, BackColor = bg,
                Padding = new Padding(4),
            };

            flp4.Controls.Add(MakeLbl("— Fill —", true));
            var rowFill = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnFillColor = MakeBtn("■ Fill Colour", 100, 26);
            btnFillColor.BackColor = Color.White;
            btnFillColor.ForeColor = Color.Black;
            btnFillColor.Click += btnFillColor_Click;
            btnNoFill = MakeBtn("No Fill", 70, 26);
            btnNoFill.Click += btnNoFill_Click;
            rowFill.Controls.AddRange(new Control[] { btnFillColor, btnNoFill });
            flp4.Controls.Add(rowFill);

            flp4.Controls.Add(MakeLbl("— Brand Swatches —", true));
            pnlFillSwatches = MakeSwatchPanel();
            flp4.Controls.Add(pnlFillSwatches);

            var rowSwatchEdit = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnEditSwatches = MakeBtn("Edit Swatches", 120, 24);
            btnEditSwatches.Click += btnEditSwatches_Click;
            rowSwatchEdit.Controls.Add(btnEditSwatches);
            flp4.Controls.Add(rowSwatchEdit);

            flp4.Controls.Add(MakeLbl("— Transparency —", true));
            var rowTrans = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            trkTransparency = new TrackBar
            {
                Minimum = 0, Maximum = 100, Value = 0, TickFrequency = 10,
                Width = 150, Height = 30, Margin = new Padding(2),
            };
            trkTransparency.Scroll += trkTransparency_Scroll;
            lblTransparencyValue = MakeLbl("0%");
            rowTrans.Controls.Add(trkTransparency);
            rowTrans.Controls.Add(lblTransparencyValue);
            flp4.Controls.Add(rowTrans);

            flp4.Controls.Add(MakeLbl("— Outline —", true));
            var rowOutline = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnOutlineColor = MakeBtn("■ Outline", 80, 26);
            btnOutlineColor.BackColor = Color.Black;
            btnOutlineColor.ForeColor = Color.White;
            btnOutlineColor.Click += btnOutlineColor_Click;
            rowOutline.Controls.Add(btnOutlineColor);
            rowOutline.Controls.Add(MakeLbl("pt:"));
            txtOutlineWidth = MakeTxt(40, "1");
            rowOutline.Controls.Add(txtOutlineWidth);
            btnApplyOutline = MakeBtn("Apply", 50, 26);
            btnApplyOutline.Click += btnApplyOutline_Click;
            btnNoOutline = MakeBtn("None", 50, 26);
            btnNoOutline.Click += btnNoOutline_Click;
            rowOutline.Controls.AddRange(new Control[] { btnApplyOutline, btnNoOutline });
            flp4.Controls.Add(rowOutline);

            flp4.Controls.Add(MakeLbl("— Outline Swatches —", true));
            pnlOutlineSwatches = MakeSwatchPanel();
            flp4.Controls.Add(pnlOutlineSwatches);

            tpgFill.Controls.Add(flp4);

            // ════════════════════════════════════════════════════════════════════
            //  TAB 5: Shadow & Quick
            // ════════════════════════════════════════════════════════════════════
            var flp5 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false, BackColor = bg,
                Padding = new Padding(4),
            };

            flp5.Controls.Add(MakeLbl("— Shadow Presets —", true));
            var rowShad1 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnShadowSoft        = MakeBtn("Soft",        60);
            btnShadowHard        = MakeBtn("Hard",        60);
            btnShadowBottom      = MakeBtn("Bottom",      60);
            btnShadowPerspective = MakeBtn("Persp.",      60);
            btnShadowRemove      = MakeBtn("Remove",      60);
            btnShadowSoft.Click        += btnShadowSoft_Click;
            btnShadowHard.Click        += btnShadowHard_Click;
            btnShadowBottom.Click      += btnShadowBottom_Click;
            btnShadowPerspective.Click += btnShadowPerspective_Click;
            btnShadowRemove.Click      += btnShadowRemove_Click;
            btnShadowRemove.ForeColor   = Color.Tomato;
            rowShad1.Controls.AddRange(new Control[] { btnShadowSoft, btnShadowHard, btnShadowBottom, btnShadowPerspective, btnShadowRemove });
            flp5.Controls.Add(rowShad1);

            flp5.Controls.Add(MakeLbl("— Quick Actions —", true));
            var rowQuick = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnDuplicate = MakeBtn("Duplicate", 85, 28);
            btnSave      = MakeBtn("Save",      85, 28);
            btnAbout     = MakeBtn("About",     85, 28);
            btnDuplicate.Click += btnDuplicate_Click;
            btnSave.Click      += btnSave_Click;
            btnAbout.Click     += btnAbout_Click;
            btnSave.BackColor   = accent;
            btnSave.ForeColor   = Color.White;
            rowQuick.Controls.AddRange(new Control[] { btnDuplicate, btnSave, btnAbout });
            flp5.Controls.Add(rowQuick);

            tpgShadow.Controls.Add(flp5);

            // ════════════════════════════════════════════════════════════════════
            //  ROOT CONTROL
            // ════════════════════════════════════════════════════════════════════
            BackColor = bg;
            Width = 290;

            Controls.Add(tabMain);
            Controls.Add(pnlHeader);
            Controls.Add(pnlFooter);
        }
    }
}
