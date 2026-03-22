using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace PPTToolbox
{
    /// <summary>
    /// Persistent swatch library. Swatches are saved to %APPDATA%\PPTToolbox\swatches.json.
    /// Defaults to BrandingConfig.DefaultSwatches on first run.
    /// </summary>
    internal static class SwatchStore
    {
        private const int MaxSwatches = 12;

        private static readonly string StorePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PPTToolbox", "swatches.json");

        private static List<Color> _swatches;

        public static IReadOnlyList<Color> Swatches
        {
            get
            {
                if (_swatches == null) Load();
                return _swatches.AsReadOnly();
            }
        }

        public static void SetSwatch(int index, Color color)
        {
            if (_swatches == null) Load();
            if (index >= 0 && index < _swatches.Count)
            {
                _swatches[index] = color;
                Save();
            }
        }

        // ── Private helpers ──────────────────────────────────────────────────────
        private static void Load()
        {
            _swatches = new List<Color>();
            try
            {
                if (File.Exists(StorePath))
                {
                    string json = File.ReadAllText(StorePath);
                    foreach (var hex in ParseJsonArray(json))
                    {
                        try { _swatches.Add(ColorTranslator.FromHtml(hex)); }
                        catch { }
                    }
                }
            }
            catch { }

            // Fill any missing slots from the defaults
            var defaults = BrandingConfig.DefaultSwatches;
            while (_swatches.Count < MaxSwatches)
            {
                int idx = _swatches.Count;
                _swatches.Add(idx < defaults.Length ? defaults[idx] : Color.White);
            }

            // Trim to max
            if (_swatches.Count > MaxSwatches)
                _swatches.RemoveRange(MaxSwatches, _swatches.Count - MaxSwatches);
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(StorePath));
                var parts = new List<string>();
                foreach (var c in _swatches)
                    parts.Add("\"" + ColorTranslator.ToHtml(c) + "\"");
                File.WriteAllText(StorePath, "[" + string.Join(",", parts) + "]");
            }
            catch { }
        }

        private static List<string> ParseJsonArray(string json)
        {
            var result = new List<string>();
            json = json.Trim();
            if (!json.StartsWith("[") || !json.EndsWith("]")) return result;
            json = json.Substring(1, json.Length - 2);
            foreach (var part in json.Split(','))
            {
                string s = part.Trim().Trim('"');
                if (!string.IsNullOrEmpty(s)) result.Add(s);
            }
            return result;
        }
    }
}
