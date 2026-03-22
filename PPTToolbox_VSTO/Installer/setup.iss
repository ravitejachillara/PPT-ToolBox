; ============================================================
;  PPT Toolbox — Inno Setup Script
;  Builds: PPTToolbox_Setup.exe
;  Requires: Inno Setup 6  (https://jrsoftware.org/isinfo.php)
;  No admin rights required (user-level install).
; ============================================================

#define MyAppName      "PPT Toolbox"
#define MyAppVersion   "1.0.0"
#define MyAppPublisher "Ravi Teja Chillara"
#define MyAppURL       ""
#define MyAppExeName   "PPTToolbox.dll"
#define BuildDir       "..\PPTToolbox\bin\Release"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={userappdata}\PPTToolbox
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=..\dist
OutputBaseFilename=PPTToolbox_Setup
SetupIconFile=..\PPTToolbox\Resources\app_icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\PPTToolbox.dll
; Show publisher banner (replace with actual logo path when available)
; WizardImageFile=..\assets\installer_banner.bmp

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WelcomeLabel1=Welcome to the [name] Setup Wizard
WelcomeLabel2=This will install [name/ver] on your computer.%n%nMade with [heart] by Ravi Teja Chillara

[Files]
; Main add-in DLL and manifest
Source: "{#BuildDir}\PPTToolbox.dll";       DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\PPTToolbox.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#BuildDir}\PPTToolbox.vsto";      DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

; All supporting DLLs in the Release folder
Source: "{#BuildDir}\*.dll";                DestDir: "{app}"; Flags: ignoreversion recursesubdirs skipifsourcedoesntexist; Excludes: "PPTToolbox.dll"
Source: "{#BuildDir}\app.config";           DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

; Resources
Source: "..\PPTToolbox\Resources\*";        DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs skipifsourcedoesntexist

[Registry]
; Register the VSTO add-in with PowerPoint (HKCU — no admin needed)
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox"; ValueType: string;  ValueName: "Description";  ValueData: "PPT Toolbox - Professional formatting add-in"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox"; ValueType: string;  ValueName: "FriendlyName"; ValueData: "PPT Tools"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox"; ValueType: dword;   ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox"; ValueType: string;  ValueName: "Manifest";     ValueData: "{app}\PPTToolbox.vsto|vstolocal"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\PPTToolbox.dll"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[UninstallRun]
; Unload add-in before uninstall
Filename: "taskkill"; Parameters: "/F /IM POWERPNT.EXE"; Flags: runhidden waituntilterminated; RunOnceId: "KillPPT"

[Code]
// Show a final message with the watermark
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssDone then
    MsgBox(
      'PPT Toolbox installed successfully!' + #13#10 + #13#10 +
      'Open PowerPoint — you will see the "PPT Tools" tab in the ribbon.' + #13#10 + #13#10 +
      'Made with [heart] by Ravi Teja Chillara',
      mbInformation, MB_OK
    );
end;
