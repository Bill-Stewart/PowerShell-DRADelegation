#define ModuleName "DRADelegation"
#define AppName ModuleName + " PowerShell Module"
#define AppPublisher "Bill Stewart"
#define AppVersion "1.0"
#define InstallPath "WindowsPowerShell\Modules\" + ModuleName
#define IconFilename "NetIQ.ico"
#define SetupCompany "Bill Stewart (bstewart@iname.com)"
#define SetupVersion "1.0.0.0"

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
Function GetWindowsPowerShellMajorVersion(): Integer;
  Var
    RootPath, VersionString: String;
    SubkeyNames: TArrayOfString;
    HighestPSVersion, I, PSVersion: Integer;
  Begin
  Result := 0;
  RootPath := 'SOFTWARE\Microsoft\PowerShell';
  If Not RegGetSubkeyNames(HKEY_LOCAL_MACHINE, RootPath, SubkeyNames) Then
    Exit;
  HighestPSVersion := 0;
  For I := 0 To GetArrayLength(SubkeyNames) - 1 Do
    Begin
    If RegQueryStringValue(HKEY_LOCAL_MACHINE, RootPath + '\' + SubkeyNames[I] + '\PowerShellEngine', 'PowerShellVersion', VersionString) Then
      Begin
      PSVersion := StrToIntDef(Copy(VersionString, 0, 1), 0);
      If PSVersion > HighestPSVersion Then
        HighestPSVersion := PSVersion;
      End;
    End;
  Result := HighestPSVersion;
  End;

Function BuildPath(Part1, Part2: String): String;
  Begin
  If Part1[Length(Part1)] <> '\' Then
    Part1 := Part1 + '\';
  Result := Part1 + Part2;
  End;

Function GetEAPath(): String;
  Var
    RootPath, InstallDir, EAPath: String;
  Begin
  Result := '';
  RootPath := 'SOFTWARE\WOW6432Node\Mission Critical Software\OnePoint\Administration';
  If Not RegQueryStringValue(HKEY_LOCAL_MACHINE, RootPath, 'InstallDir', InstallDir) Then
    Begin
    RootPath := 'SOFTWARE\Mission Critical Software\OnePoint\Administration';
    If Not RegQueryStringValue(HKEY_LOCAL_MACHINE, RootPath, 'InstallDir', InstallDir) Then
      Exit;
    End;
  EAPath := BuildPath(InstallDir, 'EA.exe');
  If FileExists(EAPath) Then Result := EAPath;
  End;

Function InitializeSetup(): Boolean;
  Var
    PSMajorVersion: Integer;
    EAPath: String;
  Begin
  PSMajorVersion := GetWindowsPowerShellMajorVersion();
  Result := PSMajorVersion >= 3;
  If Not Result Then
    Begin
    Log('FATAL: Setup cannot continue because Windows PowerShell version 3.0 or later is required.');
    If Not WizardSilent() Then
      Begin
      MsgBox('Setup cannot continue because Windows PowerShell version 3.0 or later is required.'
        + #13#10#13#10 + 'Setup will now exit.', mbCriticalError, MB_OK);
      End;
    Exit;
    End;
  Log('Windows PowerShell major version ' + IntToStr(PSMajorVersion) + ' detected');
  EAPath := GetEAPath();
  Result := EAPath <> '';
  If Not Result Then
    Begin
    Log('FATAL: Setup cannot continue because the DRA EA.exe tool is not installed.');
    If Not WizardSilent() Then
      Begin
      MsgBox('Setup cannot continue because the DRA EA.exe tool is not installed.'
        + #13#10#13#10 + 'Setup will now exit.', mbCriticalError, MB_OK);
      End;
    Exit;
    End;
  Log('EA.exe path: ' + EAPath);
  End;

Function GetInstallDir(Param: String): String;
  Begin
  If Is64BitInstallMode() Then
    Result := ExpandConstant('{commonpf64}\{#InstallPath}')
  Else
    Result := ExpandConstant('{commonpf32}\{#InstallPath}');
  End;
