# Changelog

All notable changes to PPT ToolBox are documented here.

---

## [1.0.0] — 2026-03-22

### Initial release

**Ribbon**
- Adds a single **PPT Tools** toggle button to PowerPoint's Home tab (own group, custom icon)
- Button reflects task pane state (pressed = open, released = closed)

**Task Pane — Arrange tab**
- Align to slide: Left, Centre H, Right, Top, Middle V, Bottom
- Distribute: Horizontal, Vertical (3+ shapes)
- Z-order: Bring Forward, Send Backward, Bring to Front, Send to Back
- Match size: Match Width, Match Height, Match Both (to first selected shape)
- Exact size and position inputs (cm), Read from selection

**Task Pane — Font tab**
- Font family dropdown (11 curated typefaces)
- Font size dropdown + free-type (pt)
- Bold, Italic, Underline toggles
- Font colour picker
- 12-slot colour swatch row (shared with Fill and Outline)
- Apply All button (applies family + size in one click)

**Task Pane — Para tab**
- Paragraph alignment: Left, Centre, Right, Justify
- Line spacing multiplier input
- Space Before / Space After inputs (pt)

**Task Pane — Fill tab**
- Fill colour picker, No Fill button
- 12-slot editable brand swatch palette
- Edit Swatches mode — click any swatch to replace its colour
- Transparency slider (0–100 %)
- Outline colour picker, width input (pt), Apply, No Outline
- 12-slot swatch row for outline colour (shared palette)

**Task Pane — Shadow tab**
- Presets: Soft, Hard, Bottom, Perspective
- Remove shadow

**Swatch persistence**
- Swatches saved to `%APPDATA%\PPTToolbox\swatches.json`
- Survives PowerPoint restarts and Windows reboots
- Delete the file to reset to defaults

**Footer**
- Company logo (left, max 20 px height) loaded from `Resources\company_logo.png`
- Graceful fallback to watermark text if logo is missing

**Installer**
- User-level install via Inno Setup 6 — no admin rights required
- Registers add-in under `HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox`
- Silent install flag: `/SILENT /SUPPRESSMSGBOXES /NORESTART`
- Includes uninstaller

**Xcopy deploy**
- `xcopy_deploy.bat` for no-installer option (run as target user)

---

## Roadmap

Ideas being considered for future releases — not committed.

- [ ] Keyboard shortcut to toggle the task pane
- [ ] "Reset to defaults" button for the swatch palette
- [ ] Per-presentation swatch profiles
- [ ] Slide thumbnail export (PNG/JPG) from Quick tab
- [ ] Font pairing presets (heading + body combos)
- [ ] Dark / light panel theme toggle
