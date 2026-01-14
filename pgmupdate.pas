unit PgmUpdate;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  SysUtils,
  FileUtil,
  Forms,
  Controls,
  Graphics,
  Dialogs,
  StdCtrls,
  ComCtrls,
  ExtCtrls,
  Zipper,
  uhttpdownloader;

type

  { TfrmPgmUpdate }

  TfrmPgmUpdate = class(TForm, IProgress)
    btnClose: TButton;
    edtUrl: TEdit;
    labelUrl: TLabel;
    label1: TLabel;
    label2: TLabel;
    Memo1: TMemo;
    ProgressBarDownload: TProgressBar;
    ProgressBarUnZip: TProgressBar;
    Timer1: TTimer;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { private declarations }
    procedure UnZipperProgress(Sender : TObject; const Pct: double);
    procedure StartOfFile(Sender : TObject; Const AFileName : String);
    procedure ProgressNotification(sText: String; CurrentProgress : integer; MaxProgress : integer);
    procedure Logger(sText: String);
  public
    { public declarations }
    URL : String;
    FileName : String;
  end;

var
  frmPgmUpdate: TfrmPgmUpdate;

implementation

uses
  LazUTF8,
  global,
  help;

{$R *.lfm}

{ TfrmPgmUpdate }

procedure TfrmPgmUpdate.Logger(sText: String);

begin
  memo1.Lines.Add(sText);
  myDebugLN(sText);
end;

procedure TfrmPgmUpdate.StartOfFile(Sender : TObject; Const AFileName : String);

begin
  Logger('Extract: '+AFileName);
end;

procedure TfrmPgmUpdate.UnZipperProgress(Sender : TObject; const Pct: double);

begin
  ProgressBarUnZip.Position:=round(Pct);
  ProgressBarUnZip.Refresh;
end;

procedure TfrmPgmUpdate.ProgressNotification(sText: String; CurrentProgress: integer; MaxProgress: integer);

begin
  if (MaxProgress <> -1) then ProgressBarDownload.Max:= MaxProgress;
  ProgressBarDownload.Position := CurrentProgress;
  ProgressBarDownload.Refresh;
end;

procedure TfrmPgmUpdate.FormShow(Sender: TObject);

begin
  Memo1.Clear;
  ProgressBarUnZip.Position    := 0;
  ProgressBarDownload.Position := 0;
  btnClose.Enabled             := false;
  edtUrl.Text                  := URL;
  Timer1.Enabled               := true;
end;

procedure TfrmPgmUpdate.Timer1Timer(Sender: TObject);

var
  downloader     : THttpDownloader;
  success        : boolean;
  UnZipper       : TUnZipper;
  i              : integer;
  AppName        : String;
  CurrentFileName: String;
  OldFileName    : String;
  EXE_Found      : boolean;
  sNewFiles      : TStringlist;
  nOrgFileSize   : Int64;
  dtOrgFileDate  : TDateTime;

begin
  Timer1.Enabled := false;
  EXE_Found      := false;
  AppName        := ExtractFileName(ParamStr(0));
  sNewFiles      := TStringlist.Create;

  //Idee und Unit uhttpdownloader from http://andydunkel.net/lazarus/delphi/2015/09/09/lazarus_synapse_progress.html
  Logger('Starte Download von '+URL);
  Try
    downloader := THttpDownloader.Create();
    try
      success    := downloader.DownloadHTTP(edtUrl.Text, FileName, Self);
      if Success
        then Logger('Download erfolgreich')
        else Logger('Fehler beim Download: '+downloader.sMessage);
    finally
      downloader.Free;
    end;
  except
    on E: Exception do Logger('Es ist folgender Fehler aufgetreten: '+e.Message);
  end;

  If Success
    then
      begin
        try
          UnZipper := TUnZipper.Create;
          try
            UnZipper.OnProgress := @UnZipperProgress;
            UnZipper.OnStartFile:= @StartOfFile;

            UnZipper.FileName   := UTF8ToSys(FileName);    // Name der Zip-Datei, die entpackt werden soll
            //UnZipper.OutputPath := '.';                  // Name des Ziel-Verzeichnisses

            UnZipper.Examine;

            for i := 0 to UnZipper.Entries.Count-1 do
              begin
                CurrentFileName := UnZipper.Entries.Entries[i].ArchiveFileName;
                //Logger('Zip content: '+CurrentFileName+' '+DateTimeToStr(UnZipper.Entries.Entries[i].DateTime)+' '+Inttostr(UnZipper.Entries.Entries[i].Size));

                OldFileName := CurrentFileName + '.OLD';

                //Backup erstellen
                if fileexists(CurrentFileName)
                  then
                    begin
                      if GetFileInfo(CurrentFileName, dtOrgFileDate, nOrgFileSize)
                        then
                          begin
                            if (Abs(UnZipper.Entries.Entries[i].DateTime-dtOrgFileDate)*86400 > 3) or  //Ein paar Sek. Zeitversatz erlauben
                               (UnZipper.Entries.Entries[i].Size <> nOrgFileSize)                      //Größe wird genau geprüft
                              then
                                begin
                                  //Aufräumen
                                  if fileexists(OldFileName) then
                                    begin
                                      Logger('Lösche: '+OldFileName);
                                      sysutils.deletefile(OldFileName);
                                    end;
                                  Logger('Umbenennen von : '+CurrentFileName+' nach ' +OldFileName);
                                  renamefile(CurrentFileName, OldFileName);
                                end
                              else
                                begin
                                  //Datei ist gleich --> Kein Backup, nicht entpacken
                                  Logger('Überspringen von : '+CurrentFileName+'. Datei ist unverändert');
                                  CurrentFileName := '';    //Aus Updateliste entfernen
                                end;
                          end;
                    end;

                //UnZipListe erstellen
                if CurrentFileName <> '' then sNewFiles.Add(CurrentFileName);

                if UpperCase(CurrentFileName) = UpperCase(AppName) then EXE_Found := true;
              end;

            UnZipper.UnZipFiles(sNewFiles);
            Logger('Alle Dateien entpackt.');
            if EXE_Found then Logger('Bitte '+AppName+' neu starten');
            try
              help.WriteIniVal(sIniFile, 'Programm', 'CleanUpRequired', 'true');
            except

            end;
          finally
            FreeAndNil(UnZipper);
          end;
        except
          on E: Exception do Logger('Es ist folgender Fehler aufgetreten: '+e.Message);
        end;
      end;
  sNewFiles.Free;
  btnClose.Enabled := true;
end;

end.

