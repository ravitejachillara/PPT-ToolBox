using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office     = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTToolbox
{
    /// <summary>
    /// All PowerPoint COM formatting operations.
    /// Every public method is safe to call from ribbon callbacks.
    /// </summary>
    internal static class FormattingEngine
    {
        // ── Access helpers ───────────────────────────────────────────────────────
        private static PowerPoint.Application App =>
            Globals.ThisAddIn.PPTApp;

        private static PowerPoint.Selection Selection =>
            App.ActiveWindow.Selection;

        /// <summary>Returns the selected ShapeRange, or null if nothing useful is selected.</summary>
        private static PowerPoint.ShapeRange Shapes()
        {
            try
            {
                var sel = Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    return sel.ShapeRange;
            }
            catch { }
            return null;
        }

        private static PowerPoint.Shape FirstShape()
        {
            var sr = Shapes();
            return sr != null && sr.Count > 0 ? sr[1] : null;
        }

        // ── Unit conversion ──────────────────────────────────────────────────────
        public static float CmToPoints(float cm)    => cm * 28.3465f;
        public static float PointsToCm(float pts)   => pts / 28.3465f;
        private static int  ToRgb(Color c)          => ColorTranslator.ToOle(c) & 0xFFFFFF;

        // ════════════════════════════════════════════════════════════════════════
        //  ARRANGE — Align
        // ════════════════════════════════════════════════════════════════════════
        public static void AlignLeft()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Align(Office.MsoAlignCmd.msoAlignLefts, Office.MsoTriState.msoTrue);
        }

        public static void AlignCenterH()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoTrue);
        }

        public static void AlignRight()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoTrue);
        }

        public static void AlignTop()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoTrue);
        }

        public static void AlignMiddleV()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoTrue);
        }

        public static void AlignBottom()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoTrue);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  ARRANGE — Distribute
        // ════════════════════════════════════════════════════════════════════════
        public static void DistributeH()
        {
            var sr = Shapes(); if (sr == null || sr.Count < 3) { Warn("Select 3 or more shapes to distribute."); return; }
            sr.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoTrue);
        }

        public static void DistributeV()
        {
            var sr = Shapes(); if (sr == null || sr.Count < 3) { Warn("Select 3 or more shapes to distribute."); return; }
            sr.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoTrue);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  ARRANGE — Z-Order
        // ════════════════════════════════════════════════════════════════════════
        public static void BringForward()  => ZOrder(Office.MsoZOrderCmd.msoBringForward);
        public static void SendBackward()  => ZOrder(Office.MsoZOrderCmd.msoSendBackward);
        public static void BringToFront()  => ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        public static void SendToBack()    => ZOrder(Office.MsoZOrderCmd.msoSendToBack);

        private static void ZOrder(Office.MsoZOrderCmd cmd)
        {
            var sr = Shapes(); if (sr == null) return;
            foreach (PowerPoint.Shape s in sr) s.ZOrder(cmd);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  SIZE — Match
        // ════════════════════════════════════════════════════════════════════════
        public static void MatchWidth()
        {
            var sr = Shapes(); if (sr == null || sr.Count < 2) { Warn("Select 2 or more shapes."); return; }
            float w = sr[1].Width;
            for (int i = 2; i <= sr.Count; i++) sr[i].Width = w;
        }

        public static void MatchHeight()
        {
            var sr = Shapes(); if (sr == null || sr.Count < 2) { Warn("Select 2 or more shapes."); return; }
            float h = sr[1].Height;
            for (int i = 2; i <= sr.Count; i++) sr[i].Height = h;
        }

        public static void MatchBoth()
        {
            var sr = Shapes(); if (sr == null || sr.Count < 2) { Warn("Select 2 or more shapes."); return; }
            float w = sr[1].Width, h = sr[1].Height;
            for (int i = 2; i <= sr.Count; i++) { sr[i].Width = w; sr[i].Height = h; }
        }

        // ════════════════════════════════════════════════════════════════════════
        //  SIZE — Exact W / H / X / Y  (called from task pane)
        // ════════════════════════════════════════════════════════════════════════
        public static void SetWidth(float widthCm)
        {
            var sr = Shapes(); if (sr == null) return;
            float pts = CmToPoints(widthCm);
            foreach (PowerPoint.Shape s in sr) s.Width = pts;
        }

        public static void SetHeight(float heightCm)
        {
            var sr = Shapes(); if (sr == null) return;
            float pts = CmToPoints(heightCm);
            foreach (PowerPoint.Shape s in sr) s.Height = pts;
        }

        public static void SetPosition(float xCm, float yCm)
        {
            var sr = Shapes(); if (sr == null) return;
            float xPts = CmToPoints(xCm), yPts = CmToPoints(yCm);
            foreach (PowerPoint.Shape s in sr) { s.Left = xPts; s.Top = yPts; }
        }

        /// <summary>Returns (widthCm, heightCm, leftCm, topCm) of first selected shape.</summary>
        public static (float w, float h, float x, float y) GetShapeGeometry()
        {
            var s = FirstShape();
            if (s == null) return (0, 0, 0, 0);
            return (PointsToCm(s.Width), PointsToCm(s.Height),
                    PointsToCm(s.Left),  PointsToCm(s.Top));
        }

        // ════════════════════════════════════════════════════════════════════════
        //  FONT
        // ════════════════════════════════════════════════════════════════════════
        public static void SetFontFamily(string fontName)
        {
            ForEachTextRange(tr => tr.Font.Name = fontName);
        }

        public static void SetFontSize(float sizePt)
        {
            ForEachTextRange(tr => tr.Font.Size = sizePt);
        }

        public static void ToggleBold()
        {
            ForEachTextRange(tr =>
                tr.Font.Bold = (tr.Font.Bold == Office.MsoTriState.msoTrue)
                    ? Office.MsoTriState.msoFalse
                    : Office.MsoTriState.msoTrue);
        }

        public static void ToggleItalic()
        {
            ForEachTextRange(tr =>
                tr.Font.Italic = (tr.Font.Italic == Office.MsoTriState.msoTrue)
                    ? Office.MsoTriState.msoFalse
                    : Office.MsoTriState.msoTrue);
        }

        public static void ToggleUnderline()
        {
            ForEachTextRange(tr =>
                tr.Font.Underline = (tr.Font.Underline == Office.MsoTriState.msoTrue)
                    ? Office.MsoTriState.msoFalse
                    : Office.MsoTriState.msoTrue);
        }

        public static void SetFontColor(Color color)
        {
            int rgb = ToRgb(color);
            ForEachTextRange(tr => tr.Font.Color.RGB = rgb);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  PARAGRAPH
        // ════════════════════════════════════════════════════════════════════════
        public static void SetParaAlignLeft()    => SetParaAlign(PowerPoint.PpParagraphAlignment.ppAlignLeft);
        public static void SetParaAlignCenter()  => SetParaAlign(PowerPoint.PpParagraphAlignment.ppAlignCenter);
        public static void SetParaAlignRight()   => SetParaAlign(PowerPoint.PpParagraphAlignment.ppAlignRight);
        public static void SetParaAlignJustify() => SetParaAlign(PowerPoint.PpParagraphAlignment.ppAlignJustify);

        private static void SetParaAlign(PowerPoint.PpParagraphAlignment align)
        {
            ForEachTextRange(tr => tr.ParagraphFormat.Alignment = align);
        }

        public static void SetLineSpacing(float lines)
        {
            ForEachTextRange(tr =>
            {
                tr.ParagraphFormat.LineRuleWithin = Office.MsoTriState.msoTrue;
                tr.ParagraphFormat.SpaceWithin    = lines;
            });
        }

        public static void SetSpaceBefore(float pt)
        {
            ForEachTextRange(tr =>
            {
                tr.ParagraphFormat.LineRuleBefore = Office.MsoTriState.msoFalse;
                tr.ParagraphFormat.SpaceBefore    = pt;
            });
        }

        public static void SetSpaceAfter(float pt)
        {
            ForEachTextRange(tr =>
            {
                tr.ParagraphFormat.LineRuleAfter = Office.MsoTriState.msoFalse;
                tr.ParagraphFormat.SpaceAfter    = pt;
            });
        }

        // ════════════════════════════════════════════════════════════════════════
        //  FILL
        // ════════════════════════════════════════════════════════════════════════
        public static void SetFillColor(Color color)
        {
            int rgb = ToRgb(color);
            ForEachShape(s =>
            {
                s.Fill.Solid();
                s.Fill.ForeColor.RGB = rgb;
                s.Fill.Visible = Office.MsoTriState.msoTrue;
            });
        }

        public static void SetNoFill()
        {
            ForEachShape(s => s.Fill.Visible = Office.MsoTriState.msoFalse);
        }

        public static void SetFillTransparency(float percent)
        {
            float t = Math.Max(0f, Math.Min(1f, percent / 100f));
            ForEachShape(s => s.Fill.Transparency = t);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  OUTLINE (Line)
        // ════════════════════════════════════════════════════════════════════════
        public static void SetOutlineColor(Color color)
        {
            int rgb = ToRgb(color);
            ForEachShape(s =>
            {
                s.Line.Visible        = Office.MsoTriState.msoTrue;
                s.Line.ForeColor.RGB  = rgb;
            });
        }

        public static void SetOutlineWidth(float widthPt)
        {
            ForEachShape(s =>
            {
                s.Line.Visible = Office.MsoTriState.msoTrue;
                s.Line.Weight  = widthPt;
            });
        }

        public static void SetNoOutline()
        {
            ForEachShape(s => s.Line.Visible = Office.MsoTriState.msoFalse);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  SHADOW
        // ════════════════════════════════════════════════════════════════════════
        public static void ApplyShadowSoft()
        {
            ForEachShape(s =>
            {
                var sh = s.Shadow;
                sh.Visible      = Office.MsoTriState.msoTrue;
                sh.OffsetX      = 3f;
                sh.OffsetY      = 3f;
                sh.Blur         = 10f;
                sh.Transparency = 0.40f;
                sh.ForeColor.RGB = ToRgb(Color.Black);
            });
        }

        public static void ApplyShadowHard()
        {
            ForEachShape(s =>
            {
                var sh = s.Shadow;
                sh.Visible      = Office.MsoTriState.msoTrue;
                sh.OffsetX      = 4f;
                sh.OffsetY      = 4f;
                sh.Blur         = 0f;
                sh.Transparency = 0f;
                sh.ForeColor.RGB = ToRgb(Color.Black);
            });
        }

        public static void ApplyShadowBottom()
        {
            ForEachShape(s =>
            {
                var sh = s.Shadow;
                sh.Visible      = Office.MsoTriState.msoTrue;
                sh.OffsetX      = 0f;
                sh.OffsetY      = 6f;
                sh.Blur         = 8f;
                sh.Transparency = 0.35f;
                sh.ForeColor.RGB = ToRgb(Color.Black);
            });
        }

        public static void ApplyShadowPerspective()
        {
            ForEachShape(s =>
            {
                var sh = s.Shadow;
                sh.Visible      = Office.MsoTriState.msoTrue;
                sh.OffsetX      = 0f;
                sh.OffsetY      = 12f;
                sh.Blur         = 14f;
                sh.Transparency = 0.50f;
                sh.ForeColor.RGB = ToRgb(Color.Black);
            });
        }

        public static void RemoveShadow()
        {
            ForEachShape(s => s.Shadow.Visible = Office.MsoTriState.msoFalse);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  QUICK
        // ════════════════════════════════════════════════════════════════════════
        public static void DuplicateSelected()
        {
            var sr = Shapes(); if (sr == null) return;
            sr.Duplicate();
        }

        public static void SavePresentation()
        {
            try { App.ActivePresentation.Save(); }
            catch (Exception ex) { Warn("Save failed: " + ex.Message); }
        }

        // ════════════════════════════════════════════════════════════════════════
        //  Iterator helpers
        // ════════════════════════════════════════════════════════════════════════
        private static void ForEachShape(Action<PowerPoint.Shape> action)
        {
            var sr = Shapes(); if (sr == null) return;
            foreach (PowerPoint.Shape s in sr)
            {
                try { action(s); } catch { /* skip shapes that don't support this */ }
            }
        }

        private static void ForEachTextRange(Action<PowerPoint.TextRange> action)
        {
            var sr = Shapes(); if (sr == null) return;
            foreach (PowerPoint.Shape s in sr)
            {
                try
                {
                    if (s.HasTextFrame == Office.MsoTriState.msoTrue &&
                        s.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        action(s.TextFrame.TextRange);
                    }
                }
                catch { }
            }
        }

        private static void Warn(string msg) =>
            MessageBox.Show(msg, BrandingConfig.ToolName, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
