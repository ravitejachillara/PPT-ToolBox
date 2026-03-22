====================================================================
 PPT TOOLBOX — IT DEPLOYMENT GUIDE
 Version 1.0 | Made with ♥ by Ravi Teja Chillara
====================================================================

WHAT THIS IS
------------
A PowerPoint VSTO add-in that adds a "PPT Tools" ribbon tab with
formatting shortcuts: align, distribute, z-order, font, paragraph,
fill/outline, shadow, and quick actions.

Requires: Windows 10/11, Office 2016 or later, .NET Framework 4.8
(already present on all Windows 10+ machines).


OPTION 1 — INSTALLER (recommended)
------------------------------------
Run: PPTToolbox_Setup.exe
• User-level install (no admin required).
• Copies files to %APPDATA%\PPTToolbox.
• Registers add-in in HKCU registry.
• Includes uninstaller.

If Windows shows "Unknown Publisher" warning:
  → Click "More info" → "Run anyway"  (one-time, per machine)


OPTION 2 — XCOPY (no installer)
---------------------------------
Run: xcopy_deploy.bat  as the TARGET USER
• Copies files to %APPDATA%\PPTToolbox.
• Writes HKCU registry keys.
• No elevation needed.


OPTION 3 — INTUNE / SCCM SILENT DEPLOY
-----------------------------------------
Silent install command:
  PPTToolbox_Setup.exe /SILENT /SUPPRESSMSGBOXES /NORESTART

Uninstall:
  %APPDATA%\PPTToolbox\unins000.exe /SILENT

Registry keys written (HKCU — per user):
  HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox
    Description  = "PPT Toolbox"
    FriendlyName = "PPT Tools"
    LoadBehavior = 3   (load at startup)
    Manifest     = "<install_path>\PPTToolbox.vsto|vstolocal"


TROUBLESHOOTING
---------------
Add-in not appearing in ribbon:
  1. Open PowerPoint → File → Options → Add-ins
  2. Change "Manage" dropdown to "COM Add-ins" → Go
  3. Check "PPT Tools" is listed and ticked
  4. If listed as "inactive", select it and click "Remove",
     then re-run xcopy_deploy.bat

Disabled by Office:
  File → Options → Trust Center → Trust Center Settings
  → Trusted Locations → Add Location → %APPDATA%\PPTToolbox


UNINSTALL
---------
Via Control Panel → Programs → PPT Toolbox → Uninstall
  OR
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox" /f
rmdir /s /q "%APPDATA%\PPTToolbox"

====================================================================
