#define MyAppName "Trino Excel Client"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "DAChernikov"
#define MyAppExeName "trino-excel-client-gui.exe"

[Setup]
AppId={{B91F78FB-CC19-4E3E-88A2-0F2E4A3C9D21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Trino Excel Client
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\..\dist\installer
OutputBaseFilename=TrinoExcelClientSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\..\dist\trino-excel-client-gui.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\dist\trino-excel-client-cmd.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
