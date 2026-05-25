#define MyAppName "Trino Excel Client"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "DAChernikov"
#define MyAppExeName "trino-excel-client-gui.exe"
#define MyCmdExeName "trino-excel-client-cmd.exe"

[Setup]
AppId={{B91F78FB-CC19-4E3E-88A2-0F2E4A3C9D21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Trino Excel Client
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
AlwaysShowComponentsList=yes
OutputDir=..\..\install
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Types]
Name: "full"; Description: "GUI and CLI"
Name: "gui"; Description: "Only GUI"
Name: "cli"; Description: "Only CLI"
Name: "custom"; Description: "Custom"; Flags: iscustom

[Components]
Name: "gui"; Description: "GUI application"; Types: full gui custom
Name: "cli"; Description: "Command line application"; Types: full cli custom

[Files]
Source: "..\..\dist\trino-excel-client-gui.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: gui
Source: "..\..\dist\trino-excel-client-cmd.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: cli

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Components: gui
Name: "{autoprograms}\{#MyAppName} CMD"; Filename: "{app}\{#MyCmdExeName}"; Components: cli
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Components: gui

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent; Components: gui
