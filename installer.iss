; Скрипт установки для Inno Setup
; Компиляция: ISCC.exe installer.iss

#define AppName "DMTools"
#define AppVersion "1.8.32"
#define AppExe "DMTools.exe"

[Setup]
AppId={{9F4E2C1A-5B31-4C8E-9A2D-7F3B2C0A6D11}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=ООО Дельфин центр
AppPublisherURL=
AppVerName={#AppName} {#AppVersion}
DefaultDirName={localappdata}\Programs\DMTools
DisableProgramGroupPage=yes
DefaultGroupName=DMTools
OutputDir=output
OutputBaseFilename=dmtools_setup
SetupIconFile=centrdelfin.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
WizardSizePercent=110,120
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DisableWelcomePage=no
ArchitecturesInstallIn64BitMode=x64compatible
; Без прав администратора
MinVersion=10.0

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительно:"; Flags: checkedonce

[Files]
Source: "dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\updater_helper.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "zint.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "Читать.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "centrdelfin.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"; IconFilename: "{app}\centrdelfin.ico"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; IconFilename: "{app}\centrdelfin.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Запустить {#AppName}"; Flags: nowait postinstall skipifsilent