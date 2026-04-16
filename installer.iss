[Setup]
AppName=MOAS CBC Student Report Management System
AppVersion=1.0
DefaultDirName={pf}\MOAS Report Generator
DefaultGroupName=MOAS Report Generator
OutputDir=installer
OutputBaseFilename=MOAS_Report_Generator_Installer
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\MOAS_CBC_Report.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\school_report.db"; DestDir: "{app}"; Flags: ignoreversion
Source: "moas.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "moas.jpg"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs
Source: "photos\*"; DestDir: "{app}\photos"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\MOAS Report Generator"; Filename: "{app}\MOAS_CBC_Report.exe"; IconFilename: "{app}\moas.ico"
Name: "{commondesktop}\MOAS Report Generator"; Filename: "{app}\MOAS_CBC_Report.exe"; IconFilename: "{app}\moas.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\MOAS_CBC_Report.exe"; Description: "{cm:LaunchProgram,MOAS Report Generator}"; Flags: nowait postinstall skipifsilent