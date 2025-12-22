unit appsettings;

interface

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

uses
  Classes, SysUtils, Forms;

type

 { TConfigurations }

 TConfigurations = class(TObject)
 private

 public
   {other settings as fields here}
   MyDirectory: string;
   ConfigFile: string;
   constructor Create;
   destructor Destroy; override;
 end;


var
 vConfigurations: TConfigurations;

implementation

uses LazUTF8;

{ TConfigurations }

constructor TConfigurations.Create;
begin
{$ifdef win32}
  MyDirectory := ExtractFilePath(SysToUTF8(Application.EXEName));
  ConfigFile  := MyDirectory + ChangeFileExt(ExtractFileName(Application.EXEName), '.ini');
{$endif}
{$ifdef win64}
  MyDirectory := ExtractFilePath(SysToUTF8(Application.EXEName));
  ConfigFile  := MyDirectory + ChangeFileExt(ExtractFileName(Application.EXEName), '.ini');
{$endif}
{$ifdef Unix}
  ConfigFile := GetAppConfigFile(False) + '.conf';
{$endif}
 end;

destructor TConfigurations.Destroy;
begin
  inherited Destroy;
end;

initialization

 vConfigurations := TConfigurations.Create;

finalization

 FreeAndNil(vConfigurations);

end.

