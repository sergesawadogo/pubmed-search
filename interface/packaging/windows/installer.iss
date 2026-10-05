; Installateur Windows de PubMed Search (Inno Setup 6)
; Compilation : iscc /DAppVersion=1.0.0 packaging\windows\installer.iss
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={{6F1E7C2A-3B8D-4C51-9A47-2D9E5B0F8C13}
AppName=PubMed Search
AppVersion={#AppVersion}
AppPublisher=Serge Sawadogo
DefaultDirName={autopf}\PubMed Search
DefaultGroupName=PubMed Search
DisableProgramGroupPage=yes
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\..\dist
OutputBaseFilename=PubMedSearch-Setup-{#AppVersion}
SetupIconFile=..\..\resources\icon.ico
UninstallDisplayIcon={app}\PubMedSearch.exe
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\..\dist\PubMedSearch\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\PubMed Search"; Filename: "{app}\PubMedSearch.exe"
Name: "{autodesktop}\PubMed Search"; Filename: "{app}\PubMedSearch.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PubMedSearch.exe"; Description: "{cm:LaunchProgram,PubMed Search}"; Flags: nowait postinstall skipifsilent
