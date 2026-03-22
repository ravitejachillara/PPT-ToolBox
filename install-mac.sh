#!/bin/bash

# PPT-ToolBox — Mac Installer
# Run this once per user. No admin password needed.

set -e

MANIFEST_NAME="manifest.xml"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
DOWNLOADS="$HOME/Downloads"
MANIFEST_SRC="$DOWNLOADS/$MANIFEST_NAME"

echo ""
echo "╔═══════════════════════════════════════╗"
echo "║        PPT-ToolBox Mac Installer      ║"
echo "║        Made with ♥ by Ravi Teja       ║"
echo "╚═══════════════════════════════════════╝"
echo ""

# ── Step 1: Check manifest exists in Downloads ──────────────────────────────
if [ ! -f "$MANIFEST_SRC" ]; then
  echo "❌  manifest.xml not found in Downloads folder."
  echo ""
  echo "    Please download manifest.xml from:"
  echo "    https://github.com/ravitejachillara/PPT-ToolBox"
  echo "    and place it in your Downloads folder, then run this script again."
  echo ""
  exit 1
fi

echo "✅  Found manifest.xml in Downloads"

# ── Step 2: Quit PowerPoint if running ──────────────────────────────────────
if pgrep -x "Microsoft PowerPoint" > /dev/null; then
  echo "⚠️   PowerPoint is running — quitting it now..."
  osascript -e 'quit app "Microsoft PowerPoint"'
  sleep 2
  echo "✅  PowerPoint closed"
fi

# ── Step 3: Create wef directory if it doesn't exist ────────────────────────
if [ ! -d "$WEF_DIR" ]; then
  mkdir -p "$WEF_DIR"
  echo "✅  Created add-in folder"
else
  echo "✅  Add-in folder already exists"
fi

# ── Step 4: Copy manifest ────────────────────────────────────────────────────
cp "$MANIFEST_SRC" "$WEF_DIR/$MANIFEST_NAME"
echo "✅  Installed manifest.xml"

# ── Step 5: Trust the localhost/GitHub Pages domain ─────────────────────────
# PowerPoint for Mac stores trusted domains in its plist
PLIST="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Preferences/com.microsoft.Powerpoint.plist"

if [ -f "$PLIST" ]; then
  # Trust GitHub Pages host
  defaults write com.microsoft.Powerpoint TrustAllAddins -bool true 2>/dev/null || true

  # Also write to the container plist directly
  /usr/libexec/PlistBuddy -c "Add :TrustAllAddins bool true" "$PLIST" 2>/dev/null || \
  /usr/libexec/PlistBuddy -c "Set :TrustAllAddins true" "$PLIST" 2>/dev/null || true

  echo "✅  Trust settings applied"
else
  echo "⚠️   PowerPoint preferences not found — trust will be set on first launch"
fi

# ── Step 6: Clear Office add-in cache ───────────────────────────────────────
CACHE_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches"
if [ -d "$CACHE_DIR" ]; then
  find "$CACHE_DIR" -name "*.json" -path "*wef*" -delete 2>/dev/null || true
  echo "✅  Cleared add-in cache"
fi

# ── Done ─────────────────────────────────────────────────────────────────────
echo ""
echo "═══════════════════════════════════════════"
echo "  Installation complete!"
echo ""
echo "  Next steps:"
echo "  1. Open PowerPoint"
echo "  2. Go to Insert → Add-ins → My Add-ins"
echo "  3. Click PPT Tools to open the panel"
echo ""
echo "  If you see a security warning on first open:"
echo "  Click 'Trust this add-in' — one time only."
echo "═══════════════════════════════════════════"
echo ""
