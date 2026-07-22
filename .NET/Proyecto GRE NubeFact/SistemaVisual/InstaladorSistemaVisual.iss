#define MyAppName "SistemaVisual"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "FRALSETECH"
#define MyAppExeName "SistemaVisual.exe"

[Setup]
AppId={{F3BA1A01-2B29-43EC-B510-312B412EBC71}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\FRALSETECH\SistemaVisual
DefaultGroupName=FRALSETECH

OutputDir=Salida
OutputBaseFilename=InstaladorSistemaVisual
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

UninstallDisplayIcon={app}\{#MyAppExeName}
SetupLogging=yes

[Files]
Source: "Prerequisitos\ndp48-x86-x64-allos-enu.exe"; Flags: dontcopy
Source: "bin\Release\net48\SistemaVisual.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net48\SistemaVisual.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net48\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net48\actualizador.config.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\SistemaVisual"; Filename: "{app}\SistemaVisual.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\SistemaVisual"; Filename: "{app}\SistemaVisual.exe"; WorkingDir: "{app}"

[Run]
Filename: "{app}\SistemaVisual.exe"; Description: "Ejecutar Sistema Visual ahora"; Flags: nowait postinstall skipifsilent runascurrentuser; Check: PuedeEjecutarSistemaVisual

[Code]
var
  NetFrameworkRequiereReinicio: Boolean;

function TieneNetFramework48(): Boolean;
var
  Release: Cardinal;
begin
  if IsWin64 then
    Result :=
      RegQueryDWordValue(
        HKLM64,
        'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
        'Release',
        Release
      ) and (Release >= 528040)
  else
    Result :=
      RegQueryDWordValue(
        HKLM32,
        'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
        'Release',
        Release
      ) and (Release >= 528040);
end;

function MensajeErrorNetFramework(ResultCode: Integer): String;
begin
  case ResultCode of
    1602:
      Result := 'La instalacion de Microsoft .NET Framework 4.8 fue cancelada.';
    1603:
      Result := 'Microsoft .NET Framework 4.8 no pudo instalarse debido a un error fatal.';
    5100:
      Result := 'Este equipo no cumple los requisitos para instalar Microsoft .NET Framework 4.8.';
  else
    Result :=
      'Microsoft .NET Framework 4.8 no pudo instalarse. Codigo de salida: ' +
      IntToStr(ResultCode) + '.';
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  RutaInstaladorNet: String;
  ResultCode: Integer;
begin
  Result := '';

  if TieneNetFramework48() then
    Exit;

  WizardForm.StatusLabel.Caption :=
    'Instalando Microsoft .NET Framework 4.8. Espere, por favor...';

  ExtractTemporaryFile('ndp48-x86-x64-allos-enu.exe');
  RutaInstaladorNet :=
    ExpandConstant('{tmp}\ndp48-x86-x64-allos-enu.exe');

  if not Exec(
    RutaInstaladorNet,
    '/q /norestart /ChainingPackage SistemaVisual',
    '',
    SW_SHOW,
    ewWaitUntilTerminated,
    ResultCode
  ) then
  begin
    Result :=
      'No se pudo iniciar el instalador de Microsoft .NET Framework 4.8.';
    Exit;
  end;

  if ResultCode = 0 then
  begin
    if not TieneNetFramework48() then
      Result :=
        'Microsoft .NET Framework 4.8 termino sin errores, pero no pudo verificarse en el sistema.';
  end
  else if (ResultCode = 1641) or (ResultCode = 3010) then
  begin
    NeedsRestart := True;
    NetFrameworkRequiereReinicio := True;
  end
  else
    Result := MensajeErrorNetFramework(ResultCode);
end;

function PuedeEjecutarSistemaVisual(): Boolean;
begin
  Result := not NetFrameworkRequiereReinicio;
end;
