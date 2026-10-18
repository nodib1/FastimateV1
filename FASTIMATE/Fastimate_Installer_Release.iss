; Inno Setup Script for FASTIMATE (Release Build)
; Created for Dino Bratić - FASTIMATE v2.0

#define MyAppName "Fastimate"
#define MyAppVersion "2.0"
#define MyAppPublisher "Dino Bratić"
#define MyAppExeName "FASTIMATE.exe"

[Setup]
AppId={{2CCE63B9-A1B8-416F-A414-B94133DFE266}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; Prikaz i izbor lokacije instalacije u Program Files
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=no
AlwaysShowDirOnReadyPage=yes
DirExistsWarning=yes

UninstallDisplayIcon={app}\{#MyAppExeName}
ChangesAssociations=yes
DisableProgramGroupPage=yes
OutputBaseFilename=Fastimate_Setup_v2.0
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; 1. Kopiranje SVIH Release fajlova i biblioteka (FASTIMATE.exe sa ugrađenim PDF-om, PdfSharp.dll, config)
; Putanja je relativna u odnosu na ovaj .iss fajl – prije pravljenja instalacije uraditi Release build (obfuscirani exe)
Source: "bin\Release\*"; DestDir: "{app}"; Excludes: "*.pdb"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Prečica na Desktopu i u Start meniju
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Automatsko pokretanje nakon instalacije
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
