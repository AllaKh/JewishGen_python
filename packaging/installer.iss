; installer.iss - Inno Setup script for JewishGenealogySearch.
; Compiled by build_installer.ps1 via ISCC.exe. The app version is passed on the command
; line:  ISCC.exe /DMyAppVersion=1.0.0 packaging\installer.iss
; The produced setup .exe is Authenticode-signed afterwards by build_installer.ps1.

#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppName "Jewish Genealogy Search"
#define MyAppExe  "JewishGenealogySearch.exe"
#define MyAppPublisher "Alla Khananashvili"
; folder produced by PyInstaller (relative to this .iss file's parent = repo root)
#define DistDir "..\dist\JewishGenealogySearch"

[Setup]
AppId={{6F3A9C1E-1B2D-4E7A-9C44-JEWISHGENSEARCH}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; per-user install -> no admin / UAC prompt, fewer antivirus and SmartScreen issues
PrivilegesRequired=lowest
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\Output
OutputBaseFilename=JewishGenealogySearch-Setup-{#MyAppVersion}
SetupIconFile=..\config\app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExe}
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; the entire PyInstaller output folder (exe + DLLs + config/ + storage/ + ms-playwright/)
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
