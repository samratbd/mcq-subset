; ---------------------------------------------------------------------------
; MCQ Studio — Windows installer (Inno Setup script)
;
; What this installer does:
;   * Installs MCQ_Studio.exe to %ProgramFiles%\MCQ Studio
;   * Adds Start Menu shortcut + optional Desktop shortcut
;   * Creates default folders in the user's Documents:
;       Documents\MCQ Studio\OMR Sheets    (drop scanned sheets here)
;       Documents\MCQ Studio\OMR Results
;       Documents\MCQ Studio\Question Banks (drop docx/xlsx/csv banks here)
;       Documents\MCQ Studio\Shuffled Sets
;   * Provides a proper Add/Remove Programs entry with uninstaller
;
; Requirements:
;   1. Build MCQ_Studio.exe with PyInstaller first:
;        pyinstaller --noconfirm mcq_studio.spec
;      (output: dist\MCQ_Studio.exe)
;   2. Install Inno Setup 6 (free): https://jrsoftware.org/isdl.php
;   3. Open this file in Inno Setup and click "Compile"
;   4. The installer .exe lands in installer\Output\
;
; Run silently with: MCQStudio-Setup.exe /SILENT
; Uninstall silently with: unins000.exe /SILENT
; ---------------------------------------------------------------------------

#define MyAppName "MCQ Studio"
#define MyAppVersion "1.0"
#define MyAppPublisher "Retina"
#define MyAppExeName "MCQ_Studio.exe"

[Setup]
AppId={{E6A82C2D-8F7E-4F69-8D27-MCQ-STUDIO-2026}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=Output
OutputBaseFilename=MCQStudio-Setup-{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; \
    GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
; The main executable, built by PyInstaller
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; README and license bundled for reference
Source: "..\DESKTOP_README.md"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Dirs]
; Create the default working folders in the user's Documents — they
; appear as soon as MCQ Studio is installed so users have an obvious
; place to drop files
Name: "{userdocs}\MCQ Studio"
Name: "{userdocs}\MCQ Studio\OMR Sheets"
Name: "{userdocs}\MCQ Studio\OMR Results"
Name: "{userdocs}\MCQ Studio\Question Banks"
Name: "{userdocs}\MCQ Studio\Shuffled Sets"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\OMR Results folder"; Filename: "{userdocs}\MCQ Studio\OMR Results"
Name: "{group}\Question Banks folder"; Filename: "{userdocs}\MCQ Studio\Question Banks"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; \
    Description: "Launch {#MyAppName}"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove user config when uninstalling
Type: filesandordirs; Name: "{userappdata}\..\.mcq_studio"

[Code]
{ Welcome / Finish page tweaks could go here }
