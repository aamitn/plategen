; Sticker Generator Installer Script

[Setup]
AppName=Plate Generator
AppVersion=0.6
DefaultDirName={pf}\Plate Generator
DefaultGroupName=Plate Generator
UninstallDisplayIcon={app}\plategen.exe
OutputDir=output
OutputBaseFilename=PlateGeneratorSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64

; Custom wizard images
WizardSmallImageFile=logo.bmp    


; Use the main app icon (from project root)
SetupIconFile="..\plategen_icon.ico"

[Files]
; main app executable
Source: "..\dist\plategen.exe"; DestDir: "{app}"; Flags: ignoreversion

; template docx (copied to app dir)
Source: "..\liveline_logo.dwg"; DestDir: "{app}"; Flags: ignoreversion

; app icon for shortcuts
Source: "..\plategen_icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcut
Name: "{group}\Manual Generator"; Filename: "{app}\app.exe"; IconFilename: "{app}\icon.ico"

; Desktop shortcut
Name: "{commondesktop}\Manual Generator"; Filename: "{app}\mgen.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Tasks]
; Optional desktop shortcut
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
; Option to launch app after installation
Filename: "{app}\plategen.exe"; Description: "Launch Manual Generator"; Flags: nowait postinstall skipifsilent shellexec