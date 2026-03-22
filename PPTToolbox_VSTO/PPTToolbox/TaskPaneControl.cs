using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PPTToolbox
{
    /// <summary>
    /// Main task-pane UserControl.
    /// Sections: Arrange/Size | Font | Paragraph | Fill & Outline | Shadow
    /// </summary>
    public partial class TaskPaneControl : UserControl
    {
        private bool _swatchEditMode = false;

        public TaskPaneControl()
        {
            InitializeComponent();
            PopulateFontDropdown();
            PopulateFontSizeDropdown();
            RefreshAllSwatches();
            LoadLogo();
        }

        // ── Font dropdowns ───────────────────────────────────────────────────────
        private void PopulateFontDropdown()
        {
            cmbFont.Items.Clear();
            cmbFont.Items.AddRange(BrandingConfig.BrandFonts);
            if (cmbFont.Items.Count > 0) cmbFont.SelectedIndex = 0;
        }

        private void PopulateFontSizeDropdown()
        {
            cmbFontSize.Items.Clear();
            cmbFontSize.Items.AddRange(BrandingConfig.FontSizes);
            cmbFontSize.SelectedItem = "18";
        }

        // ── Swatch management ────────────────────────────────────────────────────
        private void BuildSwatchButtons(FlowLayoutPanel panel, Action<Color> applyAction)
        {
            panel.Controls.Clear();
            var swatches = SwatchStore.Swatches;
            for (int i = 0; i < swatches.Count; i++)
            {
                int idx = i;
                var b = new Button
                {
                    Width = 24, Height = 24,
                    BackColor = swatches[i],
                    FlatStyle = FlatStyle.Flat,
                    Tag = idx,
                    Margin = new Padding(2),
                };
                b.FlatAppearance.BorderColor = BrandingConfig.BorderColor;
                b.Click += (s, e) =>
                {
                    if (_swatchEditMode)
                        EditSwatch((Button)s, idx);
                    else
                        Safe(() => applyAction(SwatchStore.Swatches[idx]));
                };
                panel.Controls.Add(b);
            }
        }

        private void RefreshAllSwatches()
        {
            BuildSwatchButtons(pnlFillSwatches,    c => FormattingEngine.SetFillColor(c));
            BuildSwatchButtons(pnlFontSwatches,    c => FormattingEngine.SetFontColor(c));
            BuildSwatchButtons(pnlOutlineSwatches, c => FormattingEngine.SetOutlineColor(c));
        }

        private void EditSwatch(Button btn, int index)
        {
            using (var dlg = new ColorDialog { FullOpen = true, Color = btn.BackColor })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    SwatchStore.SetSwatch(index, dlg.Color);
                    RefreshAllSwatches();
                }
            }
        }

        private void btnEditSwatches_Click(object s, EventArgs e)
        {
            _swatchEditMode = !_swatchEditMode;
            btnEditSwatches.Text      = _swatchEditMode ? "Done Editing" : "Edit Swatches";
            btnEditSwatches.BackColor = _swatchEditMode ? BrandingConfig.Accent : BrandingConfig.SectionBg;
            btnEditSwatches.ForeColor = _swatchEditMode ? Color.White : BrandingConfig.PanelFg;
        }

        // ── Logo loading (Change 3) ───────────────────────────────────────────────
        private void LoadLogo()
        {
            try
            {
                string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string logoPath = Path.Combine(folder, "Resources", "company_logo.png");
                if (File.Exists(logoPath))
                {
                    var img = Image.FromFile(logoPath);
                    int scaledWidth = img.Height > 0 ? (int)((float)img.Width / img.Height * 20) : 0;
                    pbLogo.Image   = img;
                    pbLogo.Width   = Math.Max(1, Math.Min(scaledWidth, 60));
                    pbLogo.Height  = 20;
                    pbLogo.Visible = true;
                }
            }
            catch { /* graceful degrade: show watermark text only */ }
        }

        // ════════════════════════════════════════════════════════════════════════
        //  TAB 1 — Arrange & Size
        // ════════════════════════════════════════════════════════════════════════
        private void btnAlignLeft_Click(object s, EventArgs e)       => Safe(FormattingEngine.AlignLeft);
        private void btnAlignCenterH_Click(object s, EventArgs e)    => Safe(FormattingEngine.AlignCenterH);
        private void btnAlignRight_Click(object s, EventArgs e)      => Safe(FormattingEngine.AlignRight);
        private void btnAlignTop_Click(object s, EventArgs e)        => Safe(FormattingEngine.AlignTop);
        private void btnAlignMiddleV_Click(object s, EventArgs e)    => Safe(FormattingEngine.AlignMiddleV);
        private void btnAlignBottom_Click(object s, EventArgs e)     => Safe(FormattingEngine.AlignBottom);
        private void btnDistributeH_Click(object s, EventArgs e)     => Safe(FormattingEngine.DistributeH);
        private void btnDistributeV_Click(object s, EventArgs e)     => Safe(FormattingEngine.DistributeV);
        private void btnBringForward_Click(object s, EventArgs e)    => Safe(FormattingEngine.BringForward);
        private void btnSendBackward_Click(object s, EventArgs e)    => Safe(FormattingEngine.SendBackward);
        private void btnBringToFront_Click(object s, EventArgs e)    => Safe(FormattingEngine.BringToFront);
        private void btnSendToBack_Click(object s, EventArgs e)      => Safe(FormattingEngine.SendToBack);
        private void btnMatchWidth_Click(object s, EventArgs e)      => Safe(FormattingEngine.MatchWidth);
        private void btnMatchHeight_Click(object s, EventArgs e)     => Safe(FormattingEngine.MatchHeight);
        private void btnMatchBoth_Click(object s, EventArgs e)       => Safe(FormattingEngine.MatchBoth);

        private void btnReadGeometry_Click(object s, EventArgs e)
        {
            var (w, h, x, y) = FormattingEngine.GetShapeGeometry();
            txtWidth.Text  = w.ToString("F2", CultureInfo.InvariantCulture);
            txtHeight.Text = h.ToString("F2", CultureInfo.InvariantCulture);
            txtX.Text      = x.ToString("F2", CultureInfo.InvariantCulture);
            txtY.Text      = y.ToString("F2", CultureInfo.InvariantCulture);
        }

        private void btnApplySize_Click(object s, EventArgs e)
        {
            if (TryParseCm(txtWidth.Text,  out float w)) Safe(() => FormattingEngine.SetWidth(w));
            if (TryParseCm(txtHeight.Text, out float h)) Safe(() => FormattingEngine.SetHeight(h));
        }

        private void btnApplyPos_Click(object s, EventArgs e)
        {
            if (TryParseCm(txtX.Text, out float x) && TryParseCm(txtY.Text, out float y))
                Safe(() => FormattingEngine.SetPosition(x, y));
        }

        // ════════════════════════════════════════════════════════════════════════
        //  TAB 2 — Font
        // ════════════════════════════════════════════════════════════════════════
        private void btnApplyFont_Click(object s, EventArgs e)
        {
            string font = cmbFont.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(font))
                Safe(() => FormattingEngine.SetFontFamily(font));

            if (float.TryParse(cmbFontSize.Text, out float sz))
                Safe(() => FormattingEngine.SetFontSize(sz));
        }

        private void btnBold_Click(object s, EventArgs e)      => Safe(FormattingEngine.ToggleBold);
        private void btnItalic_Click(object s, EventArgs e)    => Safe(FormattingEngine.ToggleItalic);
        private void btnUnderline_Click(object s, EventArgs e) => Safe(FormattingEngine.ToggleUnderline);

        private void btnFontColor_Click(object s, EventArgs e)
        {
            using (var dlg = new ColorDialog { FullOpen = true, Color = btnFontColor.BackColor })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    btnFontColor.BackColor = dlg.Color;
                    Safe(() => FormattingEngine.SetFontColor(dlg.Color));
                }
            }
        }

        // ════════════════════════════════════════════════════════════════════════
        //  TAB 3 — Paragraph
        // ════════════════════════════════════════════════════════════════════════
        private void btnParaLeft_Click(object s, EventArgs e)    => Safe(FormattingEngine.SetParaAlignLeft);
        private void btnParaCenter_Click(object s, EventArgs e)  => Safe(FormattingEngine.SetParaAlignCenter);
        private void btnParaRight_Click(object s, EventArgs e)   => Safe(FormattingEngine.SetParaAlignRight);
        private void btnParaJustify_Click(object s, EventArgs e) => Safe(FormattingEngine.SetParaAlignJustify);

        private void btnApplySpacing_Click(object s, EventArgs e)
        {
            if (float.TryParse(txtLineSpacing.Text, out float ls))
                Safe(() => FormattingEngine.SetLineSpacing(ls));
            if (float.TryParse(txtSpaceBefore.Text, out float sb))
                Safe(() => FormattingEngine.SetSpaceBefore(sb));
            if (float.TryParse(txtSpaceAfter.Text, out float sa))
                Safe(() => FormattingEngine.SetSpaceAfter(sa));
        }

        // ════════════════════════════════════════════════════════════════════════
        //  TAB 4 — Fill & Outline
        // ════════════════════════════════════════════════════════════════════════
        private void btnFillColor_Click(object s, EventArgs e)
        {
            using (var dlg = new ColorDialog { FullOpen = true, Color = btnFillColor.BackColor })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    btnFillColor.BackColor = dlg.Color;
                    Safe(() => FormattingEngine.SetFillColor(dlg.Color));
                }
            }
        }

        private void btnNoFill_Click(object s, EventArgs e) => Safe(FormattingEngine.SetNoFill);

        private void trkTransparency_Scroll(object s, EventArgs e)
        {
            lblTransparencyValue.Text = trkTransparency.Value + "%";
            Safe(() => FormattingEngine.SetFillTransparency(trkTransparency.Value));
        }

        private void btnOutlineColor_Click(object s, EventArgs e)
        {
            using (var dlg = new ColorDialog { FullOpen = true, Color = btnOutlineColor.BackColor })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    btnOutlineColor.BackColor = dlg.Color;
                    Safe(() => FormattingEngine.SetOutlineColor(dlg.Color));
                }
            }
        }

        private void btnApplyOutline_Click(object s, EventArgs e)
        {
            Safe(() => FormattingEngine.SetOutlineColor(btnOutlineColor.BackColor));
            if (float.TryParse(txtOutlineWidth.Text, out float w))
                Safe(() => FormattingEngine.SetOutlineWidth(w));
        }

        private void btnNoOutline_Click(object s, EventArgs e) => Safe(FormattingEngine.SetNoOutline);

        // ════════════════════════════════════════════════════════════════════════
        //  TAB 5 — Shadow & Quick
        // ════════════════════════════════════════════════════════════════════════
        private void btnShadowSoft_Click(object s, EventArgs e)        => Safe(FormattingEngine.ApplyShadowSoft);
        private void btnShadowHard_Click(object s, EventArgs e)        => Safe(FormattingEngine.ApplyShadowHard);
        private void btnShadowBottom_Click(object s, EventArgs e)      => Safe(FormattingEngine.ApplyShadowBottom);
        private void btnShadowPerspective_Click(object s, EventArgs e) => Safe(FormattingEngine.ApplyShadowPerspective);
        private void btnShadowRemove_Click(object s, EventArgs e)      => Safe(FormattingEngine.RemoveShadow);

        private void btnDuplicate_Click(object s, EventArgs e) => Safe(FormattingEngine.DuplicateSelected);
        private void btnSave_Click(object s, EventArgs e)      => Safe(FormattingEngine.SavePresentation);

        private void btnAbout_Click(object s, EventArgs e)
        {
            using (var dlg = new AboutDialog()) dlg.ShowDialog();
        }

        // ── Utility ──────────────────────────────────────────────────────────────
        private static void Safe(Action a)
        {
            try { a(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, BrandingConfig.ToolName, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private static bool TryParseCm(string text, out float value) =>
            float.TryParse(text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }
}
