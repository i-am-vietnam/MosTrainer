; Build after a clean Release | AnyCPU rebuild of MosTrainer/MosTrainer.slnx.
; The installed tree is the runtime Release output, not the source project tree.

[Setup]
AppId={{52562afb-1928-4adc-a450-504e13f80cb1}
AppName=MOS Excel 2019
AppVersion=1.0.0
AppVerName=MOS Excel 2019 1.0.0
DefaultDirName={autopf}\MOS Excel 2019
DefaultGroupName=MOS Excel 2019
UninstallDisplayName=MOS Excel 2019
UninstallDisplayIcon={app}\MosTrainer.exe
OutputDir=Output
OutputBaseFilename=MOS_Excel_2019_Setup_v1.0.0
SetupIconFile=..\MosTrainer\Resources\MOS_Excel_2019.ico
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
; Current XML files are reference documentation, not runtime assets. Re-audit this
; exclusion if a future project package introduces an XML resource.
Source: "..\MosTrainer\bin\Release\*"; DestDir: "{app}"; Excludes: "*.pdb,*.cs,*.xml"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\MOS Excel 2019"; Filename: "{app}\MosTrainer.exe"; IconFilename: "{app}\MosTrainer.exe"
Name: "{commondesktop}\MOS Excel 2019"; Filename: "{app}\MosTrainer.exe"; IconFilename: "{app}\MosTrainer.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\MosTrainer.exe"; Description: "Launch MOS Excel 2019"; Flags: nowait postinstall skipifsilent

[Code]
const
  DotNet472MinimumRelease = 461808;

function HasRequiredDotNet(): Boolean;
var
  ReleaseValue: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM32,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release', ReleaseValue) and (ReleaseValue >= DotNet472MinimumRelease);
  if not Result and IsWin64 then
    Result := RegQueryDWordValue(HKLM64,
      'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
      'Release', ReleaseValue) and (ReleaseValue >= DotNet472MinimumRelease);
end;

function HasExcelComInView(RootKey: Integer): Boolean;
var
  ClassId: String;
begin
  Result := RegQueryStringValue(RootKey, 'Excel.Application\CLSID', '', ClassId)
    and (ClassId <> '')
    and RegKeyExists(RootKey, 'CLSID\' + ClassId + '\LocalServer32');
end;

function HasDesktopExcel(): Boolean;
begin
  Result := HasExcelComInView(HKCR32);
  if not Result and IsWin64 then
    Result := HasExcelComInView(HKCR64);
end;

function InitializeSetup(): Boolean;
begin
  Result := False;
  if not HasRequiredDotNet() then
  begin
    MsgBox('MOS Excel 2019 requires .NET Framework 4.7.2 or later. Please install it before running Setup.',
      mbCriticalError, MB_OK);
    Exit;
  end;
  if not HasDesktopExcel() then
  begin
    MsgBox('Microsoft Excel desktop is required to run MOS Excel 2019. Please install Excel before running Setup.',
      mbCriticalError, MB_OK);
    Exit;
  end;
  Result := True;
end;
