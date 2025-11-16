; -- X2ExpressInstaller.iss
; Inno Setup script for packaging X2Express

[Setup]
AppName=X2Express
AppVersion=1.0.0
DefaultDirName={autopf}\X2Express
DefaultGroupName=X2Express
DisableProgramGroupPage=no
UninstallDisplayIcon={app}\X2Express.exe
Compression=lzma
SolidCompression=yes
OutputBaseFilename=X2Express_Installer_v1.0.0
WizardStyle=modern
PrivilegesRequired=admin

; Publisher information
AppPublisher=KUNGITEDS
AppPublisherURL=https://github.com/kungpiyaphon/ExpressAutomation
AppSupportURL=https://www.linkedin.com/in/piyaphon-waharak/
AppUpdatesURL=https://github.com/kungpiyaphon/ExpressAutomation/tree/main

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\X2Express.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "express.config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "icons\app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{commonappdata}\X2Express\incoming_exports"
Name: "{commonappdata}\X2Express\incoming_exports\processed"
Name: "{commonappdata}\X2Express\excel_templates"
Name: "{commonappdata}\X2Express\excel_templates\processed"
Name: "{commonappdata}\X2Express\logs"

[Icons]
Name: "{group}\X2Express"; Filename: "{app}\X2Express.exe"; IconFilename: "{app}\app_icon.ico"
Name: "{commondesktop}\X2Express"; Filename: "{app}\X2Express.exe"; IconFilename: "{app}\app_icon.ico"

[Tasks]
Name: "startup_shortcut"; Description: "Start X2Express when I log on"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\X2Express.exe"; Description: "Launch X2Express"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\*"
