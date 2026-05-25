#define MyAppName "Trino Excel Client"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "DAChernikov"
#define MyAppExeName "trino-excel-client-gui.exe"
#define MyCmdExeName "trino-excel-client-cmd.exe"
#define MyAppIcon "..\..\assets\app_icon.ico"

[Setup]
AppId={{B91F78FB-CC19-4E3E-88A2-0F2E4A3C9D21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Trino Excel Client
DefaultGroupName={#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=yes
AlwaysShowComponentsList=yes
CloseApplications=yes
RestartApplications=no
OutputDir=..\..\install
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#MyAppIcon}
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Types]
Name: "gui"; Description: "Только GUI"
Name: "full"; Description: "GUI и CLI"

[Components]
Name: "gui"; Description: "Графический интерфейс"; Types: gui full; Flags: fixed
Name: "cli"; Description: "Консольная утилита"; Types: full

[Files]
Source: "..\..\dist\trino-excel-client-gui.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: gui
Source: "..\..\dist\trino-excel-client-cmd.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: cli

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Components: gui; Tasks: startmenuicon
Name: "{autoprograms}\{#MyAppName} CMD"; Filename: "{app}\{#MyCmdExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Components: cli; Tasks: startmenuicon
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Components: gui

[Tasks]
Name: "startmenuicon"; Description: "Создать ярлыки в Start Menu"; GroupDescription: "Дополнительные ярлыки:"; Flags: checkedonce
Name: "desktopicon"; Description: "Создать ярлык GUI на рабочем столе"; GroupDescription: "Дополнительные ярлыки:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent; Components: gui

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C taskkill /IM {#MyAppExeName} /F /T >NUL 2>&1"; Flags: runhidden
Filename: "{cmd}"; Parameters: "/C taskkill /IM {#MyCmdExeName} /F /T >NUL 2>&1"; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}\build"
Type: dirifempty; Name: "{app}"
