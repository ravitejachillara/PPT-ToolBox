using System.Drawing;

namespace PPTToolbox
{
    /// <summary>
    /// Centralised branding constants. Update these once you have final brand assets.
    /// </summary>
    internal static class BrandingConfig
    {
        // ── Identity ─────────────────────────────────────────────────────────────
        public const string ToolName    = "PPT Tools";
        public const string Version     = "1.0.0";
        public const string Watermark   = "Made with \u2665 by Ravi Teja Chillara";
        public const string RibbonTab   = "PPT Tools";

        // ── Palette ───────────────────────────────────────────────────────────────
        // Update Primary/Accent once brand logo colours are confirmed.
        public static readonly Color Primary      = Color.FromArgb(0x1A, 0x1A, 0x2E); // deep navy
        public static readonly Color Accent       = Color.FromArgb(0xE9, 0x4F, 0x37); // vivid red-orange
        public static readonly Color PanelBg      = Color.FromArgb(0x1E, 0x1E, 0x2E); // dark panel
        public static readonly Color PanelFg      = Color.FromArgb(0xE8, 0xE8, 0xF0); // near-white text
        public static readonly Color SectionBg    = Color.FromArgb(0x28, 0x28, 0x3C); // section header
        public static readonly Color InputBg      = Color.FromArgb(0x12, 0x12, 0x20);
        public static readonly Color BorderColor  = Color.FromArgb(0x44, 0x44, 0x66);

        // ── Default swatches (initial fill for editable swatch library) ──────────
        // Editable at runtime via SwatchStore; persisted to %APPDATA%\PPTToolbox\swatches.json
        public static readonly Color[] DefaultSwatches = new[]
        {
            Color.FromArgb(0x1A, 0x1A, 0x2E), // primary dark
            Color.FromArgb(0xE9, 0x4F, 0x37), // accent red-orange
            Color.FromArgb(0xFF, 0xFF, 0xFF), // white
            Color.FromArgb(0x00, 0x00, 0x00), // black
            Color.FromArgb(0x2E, 0x86, 0xAB), // steel blue
            Color.FromArgb(0xF6, 0xC9, 0x0E), // amber
            Color.FromArgb(0x3D, 0xC4, 0x7E), // teal green
            Color.FromArgb(0xA8, 0x3F, 0x9E), // purple
            Color.FromArgb(0xFF, 0x6B, 0x35), // orange
            Color.FromArgb(0x04, 0xA7, 0x77), // emerald
            Color.FromArgb(0xD6, 0x22, 0x46), // crimson
            Color.FromArgb(0xC0, 0xC0, 0xC0), // silver
        };

        // ── Font list shown in the Font dropdown ─────────────────────────────────
        public static readonly string[] BrandFonts = new[]
        {
            "Calibri", "Calibri Light", "Arial", "Arial Narrow",
            "Helvetica Neue", "Gill Sans MT", "Century Gothic",
            "Trebuchet MS", "Verdana", "Georgia", "Times New Roman",
        };

        // ── Common font sizes ─────────────────────────────────────────────────────
        public static readonly string[] FontSizes = new[]
        {
            "8","9","10","11","12","14","16","18","20","24","28","32","36","40","48","54","60","72"
        };
    }
}
