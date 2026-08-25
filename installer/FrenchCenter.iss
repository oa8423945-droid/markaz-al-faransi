#define MyAppName "المركز الفرنسي"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "المركز الفرنسي"
#define MyAppExeName "start-system.vbs"

[Setup]
AppId={{C61BBC39-97A4-469D-B43B-30C97004C074}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\FrenchCenter
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=output
OutputBaseFilename=FrenchCenter-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
SetupLogging=yes

[Files]
Source: "stage\*"; DestDir: "{app}"; Excludes: "data\main data 2.xlsx"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "stage\data\main data 2.xlsx"; DestDir: "{app}\data"; Flags: onlyifdoesntexist uninsneveruninstall

[Icons]
Name: "{autodesktop}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\{#MyAppExeName}"""; WorkingDir: "{app}"
Name: "{group}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\{#MyAppExeName}"""; WorkingDir: "{app}"
Name: "{group}\فتح ملف البيانات"; Filename: "{app}\فتح ملف البيانات.bat"; WorkingDir: "{app}"

[Run]
Filename: "{sys}\wscript.exe"; Parameters: """{app}\{#MyAppExeName}"""; Description: "تشغيل المركز الفرنسي"; Flags: nowait postinstall skipifsilent

