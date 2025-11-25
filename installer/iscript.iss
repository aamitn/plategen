; Plategen Installer Script
; Read version from appver.txt and strip the 'v' prefix
#define RawVersion Trim(FileRead(FileOpen("..\appver.txt")))
#define AppVersion Copy(RawVersion, 2)

[Setup]
AppName=Plate Generator
AppVersion={#AppVersion}
AppPublisher=Bitmutex Technologies
DefaultDirName={pf}\Plate Generator
DefaultGroupName=Plate Generator
UninstallDisplayIcon={app}\plategen.exe
OutputDir=output
OutputBaseFilename=PlateGenSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
; Custom wizard images
WizardSmallImageFile=logo.bmp    
; Use the main app icon (from project root)
SetupIconFile=".\icons\plategen_icon.ico"

[Files]
; main app executable
Source: "..\dist\plategen.exe"; DestDir: "{app}"; Flags: ignoreversion
; ups app executable
Source: "..\dist\app_ups.exe"; DestDir: "{app}"; Flags: ignoreversion
; bch app executable
Source: "..\dist\app_bch.exe"; DestDir: "{app}"; Flags: ignoreversion
; db app executable
Source: "..\dist\app_db.exe"; DestDir: "{app}"; Flags: ignoreversion
; nameplate app executable
Source: "..\dist\app_np.exe"; DestDir: "{app}"; Flags: ignoreversion
; nameplate db schema app executable
Source: "..\dist\app_np_db_schema.exe"; DestDir: "{app}"; Flags: ignoreversion
; front back sticker app executable
Source: "..\dist\app_sticker.exe"; DestDir: "{app}"; Flags: ignoreversion
; Manual Generator Ups App Executable
Source: "..\dist\app_mgen_ups.exe"; DestDir: "{app}"; Flags: ignoreversion
; Manual Generator Ups App Executable
Source: "..\dist\app_mgen_bch.exe"; DestDir: "{app}"; Flags: ignoreversion
; logo dwg (copied to app dir)
Source: "..\liveline_logo.dwg"; DestDir: "{app}"; Flags: ignoreversion
; template docx ups(copied to app dir)
Source: "..\template-mgen-ups.docx"; DestDir: "{app}"; Flags: ignoreversion
; template docx bch(copied to app dir)
Source: "..\template-mgen-bch.docx"; DestDir: "{app}"; Flags: ignoreversion
; app icon for shortcuts
Source: ".\icons\plategen_icon.ico"; DestDir: "{app}"; Flags: ignoreversion
; appver file
Source: "..\appver.txt"; DestDir: "{app}"; Flags: ignoreversion
; sticker image file
Source: "..\sticker.png"; DestDir: "{app}"; Flags: ignoreversion
; sqlite dbfile
Source: "..\db_export\nameplates.db"; DestDir: "{app}"; Flags: ignoreversion
; autocad template file
Source: "..\acadiso.dwt"; DestDir: "{app}"; Flags: ignoreversion


[Icons]
; Start Menu shortcut
Name: "{group}\Plategen"; Filename: "{app}\plategen.exe"; IconFilename: "{app}\plategen_icon.ico"
; Desktop shortcut
Name: "{commondesktop}\Plategen"; Filename: "{app}\plategen.exe"; IconFilename: "{app}\plategen_icon.ico"; Tasks: desktopicon

[Tasks]
; Optional desktop shortcut
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
; Option to launch app after installation
Filename: "{app}\plategen.exe"; Description: "Launch Plategen"; Flags: nowait postinstall skipifsilent shellexec