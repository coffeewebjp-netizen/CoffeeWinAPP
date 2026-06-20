#define MyAppName "Coffee AutoButton"
#define MyAppPublisher "Coffee AutoButton"
#define MyAppExeName "CoffeeAutoButton.exe"

#ifndef AppVersion
#define AppVersion "1.0.10"
#endif

#ifndef PublishDir
#define PublishDir "..\publish\CoffeeAutoButton"
#endif

#ifndef OutputDir
#define OutputDir "out"
#endif

[Setup]
AppId={{8E7A6F8D-F0C8-4D31-931D-8C7903F56BA1}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\CoffeeAutoButton
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename=CoffeeAutoButtonSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成する"; GroupDescription: "追加アイコン:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} を起動する"; Flags: nowait postinstall skipifsilent
