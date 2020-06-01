#define ModuleName "DRADelegation"
#define AppName ModuleName + " PowerShell Module"
#define AppPublisher "Bill Stewart"
#define AppVersion "1.7"
#define InstallPath "WindowsPowerShell\Modules\" + ModuleName
#define IconFilename "NetIQ.ico"
#define SetupCompany "Bill Stewart (bstewart@iname.com)"
#define SetupVersion "1.7.0.0"

[Setup]
AppId={{472159EE-EB83-4332-9EED-707302449A62}
AppName={#AppName}
AppPublisher={#AppPublisher}
AppVersion={#AppVersion}
ArchitecturesInstallIn64BitMode=x64
Compression=lzma2/max
DefaultDirName={code:GetInstallDir}
DisableDirPage=yes
MinVersion=6.1
OutputBaseFilename={#ModuleName}_{#AppVersion}_Setup
OutputDir=.
PrivilegesRequired=admin
SetupIconFile={#IconFilename}
SolidCompression=yes
UninstallDisplayIcon={code:GetInstallDir}\{#IconFilename}
UninstallFilesDir={code:GetInstallDir}\Uninstall
VersionInfoCompany={#SetupCompany}
VersionInfoProductVersion={#AppVersion}
VersionInfoVersion={#SetupVersion}
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardResizable=no
WizardSizePercent=150
WizardSmallImageFile={#ModuleName}_55x55.bmp
WizardStyle=modern

[Languages]
Name: english; InfoBeforeFile: "Readme.rtf"; LicenseFile: "License.rtf"; MessagesFile: "compiler:Default.isl"

[Files]
; 32-bit
Source: "{#IconFilename}";    DestDir: "{commonpf32}\{#InstallPath}"; Check: not Is64BitInstallMode
Source: "License.txt";        DestDir: "{commonpf32}\{#InstallPath}"
Source: "Readme.md";          DestDir: "{commonpf32}\{#InstallPath}"
Source: "{#ModuleName}.psd1"; DestDir: "{commonpf32}\{#InstallPath}"
Source: "{#ModuleName}.psm1"; DestDir: "{commonpf32}\{#InstallPath}"
; 64-bit
Source: "{#IconFilename}";    DestDir: "{commonpf64}\{#InstallPath}"; Check: Is64BitInstallMode
Source: "License.txt";        DestDir: "{commonpf64}\{#InstallPath}"; Check: Is64BitInstallMode
Source: "Readme.md";          DestDir: "{commonpf64}\{#InstallPath}"; Check: Is64BitInstallMode
Source: "{#ModuleName}.psd1"; DestDir: "{commonpf64}\{#InstallPath}"; Check: Is64BitInstallMode
Source: "{#ModuleName}.psm1"; DestDir: "{commonpf64}\{#InstallPath}"; Check: Is64BitInstallMode

[Code]
function GetWindowsPowerShellMajorVersion(): integer;
  var
    RootPath,VersionString: string;
    SubkeyNames: TArrayOfString;
    HighestPSVersion,I,PSVersion: integer;
  begin
  result := 0;
  RootPath := 'SOFTWARE\Microsoft\PowerShell';
  if not RegGetSubkeyNames(HKEY_LOCAL_MACHINE,RootPath,SubkeyNames) then
    exit;
  HighestPSVersion := 0;
  for I := 0 to GetArrayLength(SubkeyNames) - 1 do
    begin
    if RegQueryStringValue(HKEY_LOCAL_MACHINE,RootPath + '\' + SubkeyNames[I] + '\PowerShellEngine','PowerShellVersion',VersionString) then
      begin
      PSVersion := StrToIntDef(Copy(VersionString,0,1),0);
      if PSVersion > HighestPSVersion then
        HighestPSVersion := PSVersion;
      end;
    end;
  result := HighestPSVersion;
  end;

function IsDRAServer(): boolean;
  var
    EAServer: variant;
  begin
  result := false;
  try
    EAServer := CreateOleObject('EAServer.EAServe');
    result := true;
  except
  end; //try
  End;

function BuildPath(Part1,Part2: string): string;
  begin
  if Part1[Length(Part1)] <> '\' then
    Part1 := Part1 + '\';
  Result := Part1 + Part2;
  end;

function IsEAInstalled(): boolean;
  var
    RootPath,InstallDir,EAPath: string;
  begin
  result := false;
  RootPath := 'SOFTWARE\WOW6432Node\Mission Critical Software\OnePoint\Administration';
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,RootPath,'InstallDir',InstallDir) then
    begin
    RootPath := 'SOFTWARE\Mission Critical Software\OnePoint\Administration';
    if not RegQueryStringValue(HKEY_LOCAL_MACHINE,RootPath,'InstallDir',InstallDir) then
      exit;
    end;
  EAPath := BuildPath(InstallDir,'EA.exe');
  result := FileExists(EAPath);
  end;

Function InitializeSetup(): Boolean;
  var
    PSMajorVersion: integer;
  begin
  PSMajorVersion := GetWindowsPowerShellMajorVersion();
  result := PSMajorVersion >= 3;
  if not result then
    begin
    Log('FATAL: Setup cannot continue because Windows PowerShell version 3.0 or later is required.');
    if not WizardSilent() then
      begin
      MsgBox('Setup cannot continue because Windows PowerShell version 3.0 or later is required.'
        + #13#10#13#10 + 'Setup will now exit.',mbCriticalError,MB_OK);
      end;
    exit;
    end;
  Log('Windows PowerShell major version ' + IntToStr(PSMajorVersion) + ' detected');
  result := ISDRAServer();
  if not result then
    begin
    Log('FATAL: Setup cannot continue because the current computer is not a DRA server.');
    if not WizardSilent() then
      begin
      MsgBox('Setup cannot continue because the current computer is not a DRA server.'
        + #13#10#13#10 + 'Setup will now exit.',mbCriticalError,MB_OK);
      end;
    exit;
    end;
  result := IsEAInstalled();
  if not result then
    begin
    Log('FATAL: Setup cannot continue because the DRA command-line interface feature is not installed.');
    if not WizardSilent() then
      begin
      MsgBox('Setup cannot continue because the DRA command-line interface feature is not installed.'
        + #13#10#13#10 + 'Setup will now exit.',mbCriticalError,MB_OK);
      end;
    exit;
    end;
  Log('DRA command-line interface feature detected');
End;

function GetInstallDir(Param: string): string;
  begin
  if Is64BitInstallMode() then
    result := ExpandConstant('{commonpf64}\{#InstallPath}')
  else
    result := ExpandConstant('{commonpf32}\{#InstallPath}');
  end;
