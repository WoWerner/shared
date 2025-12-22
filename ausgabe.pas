unit Ausgabe;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  SysUtils,
  LazUTF8,
  Forms,
  Controls,
  Graphics,
  Dialogs,
  ExtCtrls,
  StdCtrls,
  Menus, Buttons,
  PrintersDlgs,
  SynMemo;

type

  { TfrmAusgabe }

  TfrmAusgabe = class(TForm)
    btnClose: TButton;
    btnOK: TButton;
    btnSave: TButton;
    btnPrint: TButton;
    btnCopy: TButton;
    mnuClear: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    PopupMenuMemo: TPopupMenu;
    PrintDialog: TPrintDialog;
    SaveDialog: TSaveDialog;
    Memo: TSynMemo;
    procedure btnCopyClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure mnuClearClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
    procedure SetDefaults(sCaption, sText, sFName, sOKButtonCaption, sCloseButtonCaption: string; bOKButtonVisible: boolean);
  end; 

var
  frmAusgabe: TfrmAusgabe;

implementation

uses
  global,
  LConvEncoding,
  help;

{$R *.lfm}

{ TfrmAusgabe }

var
  sFileName : String;

procedure TfrmAusgabe.btnCopyClick(Sender: TObject);
begin
  memo.SelectAll;
  memo.CopyToClipboard;
  memo.SelStart  := 0;
  memo.SelEnd    := 0;
end;

procedure TfrmAusgabe.btnPrintClick(Sender: TObject);
begin
  PrintText(memo.text, 'Ausgabe', PrintDialog);
end;

procedure TfrmAusgabe.btnSaveClick(Sender: TObject);

var
  sFile  : string;

begin
  if sFileName = '' then sFileName := sPrintPath+'Ausgabe.txt';
  // Set up the starting directory to be the current one
  SaveDialog.InitialDir := ExtractFilePath(sFileName);
  SaveDialog.Options    := [];
  SaveDialog.Filter     := 'Alle|*.*|Textdatei (*.txt)|*.txt|Excel (*.csv)|*.csv';

  // Select TXT files as the starting filter type
  SaveDialog.FilterIndex := 2;
  SaveDialog.FileName := sFileName;

  // Display the open file dialog
  if SaveDialog.Execute
    then
      begin
        sFile := SaveDialog.FileName;

        memo.Lines.SaveToFile(sFile);
        //Alternative in anderen Formaten speichern
        memo.Text:=UTF8toCP1252(memo.Text);
        memo.Lines.SaveToFile(ExtractFilePath(sFile)+StringReplace(ExtractFileName(sFile), '.', '_ANSI_CP1252.', []));

        //Dialog schliessen
        close;
      end;
end;

procedure TfrmAusgabe.FormShow(Sender: TObject);
begin
  memo.TopLine := 0;
  memo.SetFocus;
end;

procedure TfrmAusgabe.mnuClearClick(Sender: TObject);
begin
  memo.Clear;
end;

procedure TfrmAusgabe.SetDefaults(sCaption, sText, sFName, sOKButtonCaption, sCloseButtonCaption: string; bOKButtonVisible: boolean);
begin
  memo.Text        := sText;
  caption          := sCaption;
  sFileName        := sFName;
  btnClose.Caption := sCloseButtonCaption;
  btnOK.Caption    := sOKButtonCaption;
  btnOK.Visible    := bOKButtonVisible;
end;

end.

