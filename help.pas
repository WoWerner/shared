unit help;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  SysUtils,
  inifiles,
  printers,
  LazUTF8,
  Graphics,
  Forms,
  Controls,
  DBCtrls,
  DBGrids,
  db,
  Dialogs,
  ZDataset,
  PrintersDlgs,
  regexpr;

type
  TCharSet = set of Char;
  TMargins = record
    Left,
    Top,
    Right,
    Bottom: Double
  end;

//IniFile
function  ReadIniVal(FileName, Section, Keyword, DefVal: string; StoreDef : boolean): string;
function  ReadIniBool(FileName, Section, Keyword: string; DefVal, StoreDef : boolean): boolean;
function  ReadIniInt(FileName, Section, Keyword:String; DefVal: integer): integer;
procedure WriteIniVal(FileName, Section, Keyword, Val: string);
procedure WriteIniBool(FileName, Section, Keyword: string; b: boolean);
procedure WriteIniInt(FileName, Section, Keyword: string; Val: integer);

//Passwort
function  verschluessele(passwort : string):string;

//String
function  AppendChar(Source: String; AddChar: char; len: integer):string; //hängt zeichen an
function  ReplaceChar(int_Str:  string; cDel, cIn : Char) : string;
function  ReplaceChars(int_Str:  string; cDel: TCharSet; cIn : Char) : string;
function  DeleteChar(int_Str: string; cDel : Char) : string;
function  DeleteChars(int_Str: string; cDel : TCharSet) : string;
function  CountChar(int_Str: string; CSearch : Char) : integer;
function  Like(const AString, APattern: String): Boolean;
function  RemoveBOM(text: string): String;
function  RemoveLastCRLF(text: string): String;

function  GetCSVRecordItem(const ItemIndex     : Integer;
                           const CSVRecord     : string;
                           const WordSep       : TCharSet;
                           const WordDelimiter : Char): string;
function  ExtractWord(     const N             : Integer;
                           const S             : string;
                           const WordSep       : TCharSet): string;
function  WordPosition(    const N             : Integer;
                           const S             : string;
                           const WordSep       : TCharSet): Integer;
function  OnlyDigits(s: string): boolean;
function  XCharsOnly(s: string; const XChar: TCharSet): string;


// Konvertierung
function  myval(int_zahl_str: string) : longint;
function  RealToStr(D:Double;Stellen,Komma:Integer):String;
function  IntToCurrency(i: Longint): string;
function  CurrencyToInt(s: string; bEuroM: boolean): Longint;
Function  ZahlInString(n:Longint):String;
Function  BoolInString(b:boolean):String;

//Datum
function AgeNow(const DateOfBird: TDateTime): Integer;
function AgeAtDate(const DateOfBird, DT: TDateTime): Integer;
function year(int_datum: string) : String;  //Erlaubtes Datumsformat DD.MM.YYYY oder YYYY-MM-DD
function IsDateFormat(s : String; Date_Del: Char): Boolean;

//Programm
function GetProductVersionString: String;

//Debug
Procedure myDebugLN(s: string);
Procedure myDebugStart(s: string);
Procedure myDebugAdd(s: string);
Procedure myDebugFinish(s: string);
Procedure WriteDebug(s: string);
Procedure FlushDebug();
procedure GetPrinterMargins();
procedure LogAndShowError(s: string);
procedure LogAndShow(s: string);

//SQL
function SQL_Where_Contains(s, Feld : String): string;
function SQL_Where_IsNull(Feld : String): string;
Function SQL_Where_Add(sWhere, sAdd:string): string;
Function SQL_Where_Add_OR(sWhere, sAdd:string):string;
function SQL_UTF8UmlautReplace(Feld : String): string;
function SQL_DeleteComment(SQL_IN : String): string;
function SQL_QuotedStr(SQL_IN : String): string;
function FieldTypeToString(FT: TFieldtype): string;
function SQLiteDateFormat(d: TDateTime): string;

//Datenbank
function  ExecSQL(SQLCommand : string; MyQue: TZQuery; ShowUpdateMessage: boolean):integer;
procedure ExportQueToCSVFile(Que: TZQuery; const FileName, sTrenner, sDelimiter: String; FinishMessage, UTF8: boolean; IncludeNullEuro: boolean = true);
function  GetFirstDBFieldAsStringList(Que: TZQuery):String;
function  GetDBSum(Query: TZQuery; table, field, join, where : string):longint;
function  IsValInTable(Query: TZQuery; table, field, value : string):boolean;

//Drucken
Procedure PrintText(sDruckText, sTitel: String; PrtDlg: TPrintDialog);

//Windows
function GetComputerName: string;
function GetUserName: string;
function GetWorkArea: TRect;
function GetMaxWindowsSize: TRect;
function GetCaptionHeight: integer;
function GetMonitorCount: Integer;
function GetVirtualScreenSize: TRect;

//Files
function GetFileInfo(const AFileName: String; var FileWriteTime: TDateTime; var FileSize: Int64): Boolean;

var
  bTausendertrennung : boolean = false;
  bEuromodus         : boolean = false;
  bDebug             : boolean = true;

{******************************************************************************}
{******************************************************************************}

implementation

uses
  ZDbcIntfs,
  global,
  LConvEncoding,
  LCLIntf,
  versiontypes,
  windows,
  vinfo;

Const OutPutDataTypes = [ftUnknown,
                         ftString,
			 ftMemo,
                         ftSmallint,
                         ftInteger,
                         ftLargeint,
                         ftWord,
                         ftBoolean,
                         ftFloat,
                         ftCurrency,
                         ftDate,
                         ftTime,
                         ftDateTime,
                         ftBytes,
                         ftVarBytes,
                         ftAutoInc];

var
  sDebug : string = '';

//****************************************************************

function GetFileInfo(const AFileName: String; var FileWriteTime: TDateTime; var FileSize: Int64): Boolean;

var
  SR            : TSearchRec;
  LocalFileTime : TFileTime;
  SystemTime    : TSystemTime;

begin
  Result := False;
  if FindFirst(AFileName, faAnyFile, SR) = 0
    then
      begin
        if FileTimeToLocalFileTime(SR.FindData.ftLastWriteTime, LocalFileTime)
          then
            begin
              if FileTimeToSystemTime(LocalFileTime, SystemTime)
                 then
                   begin
                     FileWriteTime := SystemTimeToDateTime(SystemTime);
                     FileSize      := SR.FindData.nFileSizeLow or (SR.FindData.nFileSizeHigh shl 32);
                     Result := True;
                   end;
            end;
        SysUtils.FindClose(SR);
      end;
end;

function GetMonitorCount: Integer;
begin
  Result := GetSystemMetrics(SM_CMONITORS);
  myDebugLN('MonitorCount: '+inttostr(Result));
end;

function GetVirtualScreenSize: TRect;

//In Right  -->  Width
//In Bottom -->  Height

begin
  result.Left   := GetSystemMetrics(SM_XVIRTUALSCREEN);
  result.Top    := GetSystemMetrics(SM_YVIRTUALSCREEN);
  result.Right  := result.Left + GetSystemMetrics(SM_CXVIRTUALSCREEN);
  result.Bottom := result.Top  + GetSystemMetrics(SM_CYVIRTUALSCREEN);
  myDebugLN(#13#10+
            'VirtualScreenSize:'+#13#10+
            'Left  : '+inttostr(Result.Left)+#13#10+
            'Top   : '+inttostr(Result.Top)+#13#10+
            'Width : '+inttostr(Result.Right)+' (store in Right)'+#13#10+
            'Heigth: '+inttostr(Result.Bottom)+' (store in Bottom)');
end;

function GetWorkArea: TRect;
begin
  SystemParametersInfo(SPI_GETWORKAREA, 0, @Result, 0);
  myDebugLN(#13#10+
            'WorkArea:'+#13#10+
            'Left  : '+inttostr(Result.Left)+#13#10+
            'Top   : '+inttostr(Result.Top)+#13#10+
            'Width : '+inttostr(Result.Right)+' (store in Right)'+#13#10+
            'Heigth: '+inttostr(Result.Bottom)+' (store in Bottom)');
end;

function GetCaptionHeight : integer;

begin
  result := GetSystemMetrics(SM_CYCAPTION)
end;

function GetMaxWindowsSize: TRect;

//In Right  -->  Width
//In Bottom -->  Height

var
  aBorderWidth,
  aCaption     : integer;
  WORKAREA     : TRect;

begin
  WORKAREA     := GetWorkArea; //Verfügbarer Platz
  aBorderWidth := GetSystemMetrics(SM_CYSIZEFRAME);
  aCaption     := GetSystemMetrics(SM_CYCAPTION);
  Result.Left   := WORKAREA.Left;
  Result.Top    := WORKAREA.Top;
  Result.Right  := WORKAREA.Right  - WORKAREA.Left - 2*aBorderWidth;
  Result.Bottom := WORKAREA.Bottom - WORKAREA.Top  -   aBorderWidth - aCaption;
  myDebugLN(#13#10+
            'PosAndMaxWindowsSize:'+#13#10+
            'Left  : '+inttostr(Result.Left)+#13#10+
            'Top   : '+inttostr(Result.Top)+#13#10+
            'Width : '+inttostr(Result.Right)+' (store in Right)'+#13#10+
            'Heigth: '+inttostr(Result.Bottom)+' (store in Bottom)');
end;

//****************************************************************

function IsValInTable(Query: TZQuery; table, field, value : string):boolean;

begin
  Query.SQL.Text := 'select '+field+' from '+table+' where '+field+'='+value;
  Query.Open;
  result := Query.RecordCount > 0;
  Query.Close;
end;

//****************************************************************

function GetDBSum(Query: TZQuery; table, field, join, where : string):longint;

begin
  Query.SQL.Text := 'select sum('+field+') as MySum from '+table+' '+join+' where '+where;
  try
    Query.Open;
    result := Query.FieldByName('MySum').AsLongint;
    Query.Close;
  except
    //Tritt auf, wenn alte Jahre ohne Daten gewählt werden
    result := 0;
  end;
end;

//****************************************************************

procedure LogAndShowError(s: string);

begin
  LogAndShow('Error Message: '+s);
end;

//****************************************************************

procedure LogAndShow(s: string);

begin
  myDebugLN(s);
  Showmessage(s);
end;

//****************************************************************

function RemoveBOM(text: string): String;

const UTF8BOM:string=#$EF#$BB#$BF;

begin
  result := text;
  if length(text)>=3
    then
      if Copy(text,1,3)=UTF8BOM
         then
            begin
              Delete(text,1,3);
              result := text;
            end;
end;

//****************************************************************

function RemoveLastCRLF(text: string): String;

begin
  result := text;
  while (length(result) > 0 ) and (result[length(result)] in [#10, #13]) do
    result := Copy(result,1,length(result)-1);
end;

//****************************************************************

function IsDateFormat(s : String; Date_Del: Char): Boolean;

var
  RegexObj: TRegExpr;

begin
  // result := Like(s,'*?'+Date_Del+'*?'+Date_Del+'??**') and (length(s) in [6..10]);
  RegexObj := TRegExpr.Create;
  RegexObj.Expression := '^(19|20)\d\d['+Date_Del+']([1-9]|0[1-9]|1[012])['+Date_Del+']([1-9]|0[1-9]|[12][0-9]|3[01])$|'+  //2015.12.31
                         '^([1-9]|0[1-9]|[12][0-9]|3[01])['+Date_Del+']([1-9]|0[1-9]|1[012])['+Date_Del+'](19|20)\d\d$';   //31.12.2015
  result := RegexObj.Exec(s);
  RegexObj.Free;
end;

{******************************************************************************}

function Like(const AString, APattern: String): Boolean;

{ Like prüft die Übereinstimmung eines Strings mit einem Muster.
  So liefert Like('Delphi', 'D*p?i') true.
  Der Vergleich berücksichtigt Klein- und Großschreibung.
  Ist das nicht gewünscht, muss statt dessen
  Like(AnsiUpperCase(AString), AnsiUpperCase(APattern)) benutzt werden:
  Michael Winter}

var
  StringPtr, PatternPtr: PChar;
  StringRes, PatternRes: PChar;

begin
  Result    := false;
  StringPtr := PChar(AString);
  PatternPtr:= PChar(APattern);
  StringRes := nil;
  PatternRes:= nil;
  repeat
    repeat // ohne vorangegangenes "*"
      case PatternPtr^ of
        #0:  begin
               Result:=StringPtr^=#0;
               if Result or (StringRes=nil) or (PatternRes=nil) then Exit;
               StringPtr:=StringRes;
               PatternPtr:=PatternRes;
               Break;
             end;
        '*': begin
               inc(PatternPtr);
               PatternRes:=PatternPtr;
               Break;
             end;
        '?': begin
               if StringPtr^=#0 then Exit;
               inc(StringPtr);
               inc(PatternPtr);
             end;
        else begin
               if StringPtr^=#0 then Exit;
               if StringPtr^<>PatternPtr^
                 then
                   begin
                     if (StringRes=nil) or (PatternRes=nil) then Exit;
                     StringPtr:=StringRes;
                     PatternPtr:=PatternRes;
                     Break;
                   end
                 else
                   begin
                     inc(StringPtr);
                     inc(PatternPtr);
                   end;
             end;
      end; //Case
    until false;
    repeat // mit vorangegangenem "*"
      case PatternPtr^ of
        #0:  begin
               Result:=true;
               Exit;
             end;
        '*': begin
               inc(PatternPtr);
               PatternRes:=PatternPtr;
             end;
        '?': begin
               if StringPtr^=#0 then Exit;
               inc(StringPtr);
               inc(PatternPtr);
              end;
        else begin
               repeat
                 if StringPtr^=#0 then  Exit;
                 if StringPtr^=PatternPtr^ then Break;
                 inc(StringPtr);
               until false;
               inc(StringPtr);
               StringRes:=StringPtr;
               inc(PatternPtr);
               Break;
             end;
      end; //Case
    until false;
  until false;
end;

{******************************************************************************}

function GetComputerName: string;

var
  buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: Cardinal;

begin
  Size := MAX_COMPUTERNAME_LENGTH + 1;
  Windows.GetComputerName(@buffer, Size);
  Result := StrPas(buffer);
end;

{******************************************************************************}

function GetUserName: string;

var
  buffer: array[0..255] of Char;
  Size: Cardinal;

begin
  Size := 255;
  Windows.GetUserName(@buffer, Size);
  Result := StrPas(buffer);
end;

{******************************************************************************}

function GetFirstDBFieldAsStringList(Que: TZQuery):String;

begin
  Que.Open;
  result := '';
  While not Que.EOF
    do
      begin
        result := result + Que.Fields[0].AsString+#13#10;
        Que.Next;
      end;
  Que.Close;
end;

procedure ExportQueToCSVFile(Que: TZQuery; const FileName, sTrenner, sDelimiter: String; FinishMessage, UTF8: boolean; IncludeNullEuro: boolean = true);

var f: TextFile;
    i,
    j: integer;
    s: String;

begin
  Screen.Cursor := crHourglass;

  j := 0;

  //Datei anlegen
  try
    ForceDirectories(ExtractFilePath(FileName));
  except
    //Wenn ohne Path
  end;

  try
    AssignFile(F, Filename);
    Rewrite(f);

    //Headerausgabe
    for i := 1 to que.FieldCount do
      begin
        if Que.FieldDefs[i-1].DataType in OutPutDataTypes
          then
            begin
              if i > 1 then write(F, sTrenner);
              if UTF8
                then s := Que.FieldDefs[i-1].Name
                else s := UTF8toCP1252(Que.FieldDefs[i-1].Name{+'('+FieldTypeToString(Que.FieldDefs[i-1].DataType)+')'});
              if (pos(sTrenner, s) > 0) or (Que.FieldDefs[i-1].DataType in [ftString, ftMemo])
                 then s := sDelimiter+s+sDelimiter;
              Write(F, s);
            end;
      end;
    WriteLn(F, '');

 //{$DEFINE Field_DEBUG}
{$ifdef Field_DEBUG}
    //Debug Feldtypeausgabe
    for i := 1 to que.FieldCount do
      begin
        if Que.FieldDefs[i-1].DataType in OutPutDataTypes
          then
            begin
              if i > 1 then write(F, sTrenner);
              if not UTF8 then s := UTF8toCP1252(Que.FieldDefs[i-1].Name+'('+FieldTypeToString(Que.FieldDefs[i-1].DataType)+')');
              if pos(sTrenner, s) > 0 then s := sDelimiter+s+sDelimiter;
              Write(F, s);
            end;
      end;
    WriteLn(F, '');
{$endif}

    //Datenausgabe
    Que.DisableControls;
    Que.Refresh;
    Que.First;
    While not Que.EOF do
      begin
        inc(j);
        for i := 1 to que.FieldCount do
          begin
            if Que.FieldDefs[i-1].Datatype in OutPutDataTypes
              then
                begin
                  if i > 1 then write(F, sTrenner);
                  case Que.FieldDefs[i-1].DataType of
                    ftDate:     s := formatdatetime('dd.mm.yyyy',Que.Fields[i-1].AsDateTime);
                    ftFloat,
                    ftCurrency: begin
                                  if bTausendertrennung
                                    then s := format('%.2n',[Que.Fields[i-1].AsCurrency])
                                    else s := format('%.2f',[Que.Fields[i-1].AsCurrency]);
                                end;
		    ftMemo    : begin
		                  s := ReplaceChar(Que.Fields[i-1].asstring, #13, ' ');
				  s := ReplaceChar(s, #10, ' ');
				  s := StringReplace(s, '  ', ' ', [rfReplaceAll]);
		                end;
                    else
                                  s := Que.Fields[i-1].asstring;
                  end;
                  s := DeleteChars(s, [#10,#13]);
                  if not IncludeNullEuro and (s='0,00 €') then s := '';
                  if (pos(sTrenner, s) > 0) or (Que.FieldDefs[i-1].DataType in [ftString, ftMemo])
                    then s := sDelimiter+s+sDelimiter;
                  if not UTF8 then s := UTF8toCP1252(s);
                  Write(F, s);
                end;
          end;
        WriteLn(F, '');
        Que.Next;
      end;
    Que.First;
    Que.EnableControls;

    CloseFile(f); //Ausgabedatei schliessen

    Screen.Cursor := crDefault;
    if finishMessage
       then if MessageDlg(Inttostr(j)+' Zeilen exportiert in Datei '+FileName+#13+
                          'Sollen Sie angezeigt werden?', mtConfirmation, [mbYes, mbNo],0)= mrYes
          then opendocument(FileName);
  except
    on E : Exception do
      begin
        Screen.Cursor := crDefault;
        LogAndShowError('Fehler beim Schreiben der Datei: '+Filename+#13#13+E.ClassName+ ': '+ E.Message);
      end;
  end;

  try
    CloseFile(f); //Ausgabedatei schliessen
  except
    //Tritt auf, wenn die Datei nicht offen war
  end;
end;

{******************************************************************************}

Function ZahlInString(n:Longint):String;

Const
  Zahlen1 : Array[0..9] Of String = ('','zehn','zwan','drei','vier','fünf','sech','sieb','ach','neun');
  Zahlen  : Array[0..9] Of String = ('','ein','zwei','drei','vier','fünf','sechs','sieben','acht','neun');

Var
  n100,
  n10,
  n1   : Integer;
  s    : String;

  Function ZehnerUndEiner(n10,n1:Byte):String;

  Var n:Integer;

  Begin
    n:=n10*10+n1;
    Result:='';
    If n10=0 Then
      Begin
        If n1>0 Then Result:=Result+Zahlen[n1];
        //If n1=1 Then Result:=Result+'s';
      End
    Else
      Begin
        If n10=1
          Then
            Begin
              If n=11
                Then Result:=Result+'elf'
                Else If n=12
                  Then Result:=Result+'zwölf'
                  Else Result:=Result+Zahlen1[n1]+'zehn';
             End
          Else
            Begin
              Result:=Result+Zahlen[n1];
              If n1>0 Then Result:=Result+'und';
              Result:=Result+Zahlen1[n10];
              If n10<>3
                Then Result:=Result+'zig'
                Else Result:=Result+'ßig';
            End;
      End;
  End; {ZehnerUndEiner}

begin
  Result:='';
  If n=0
    Then
      Begin
        Result:='null';
      End
    else
      begin
        if n<0
          then
            begin
              Result:='minus ';
              n := Abs(n);
            end;
        If n>=1000000000 Then Begin
          s:=ZahlInString(n DIV 1000000000);
          If s='eins'
            Then Result:=Result+'einemilliarde'
            Else Result:=Result+s+'milliarden';
          n:=n MOD 1000000000;
        End;
        If n>=1000000 Then Begin
          s:=ZahlInString(n DIV 1000000);
          If s='eins'
            Then Result:=Result+'einemillion'
            Else Result:=Result+s+'millionen';
          n:=n MOD 1000000;
        End;
        If n>=1000 Then Begin
          s:=ZahlInString(n DIV 1000);
          If s='eins' Then s:='ein';
          Result:=Result+s+'tausend';
          n:=n MOD 1000;
        End;
        n100 := n Div 100;
        n    := n MOD 100;
        n10  := n Div 10;
        n1   := n Mod 10;
        If n100<>0 Then Result:=Result+Zahlen[n100]+'hundert';
        Result:=Result+ZehnerUndEiner(n10,n1);
        Result[1] := Uppercase(Result[1])[1];
      end;
end; {Basis Georg W. Seefried}

Function  BoolInString(b:boolean):String;
begin
  if b
    then result := 'TRUE'
    else result := 'FALSE';
end;

{******************************************************************************}

function CountChar(int_Str: string; CSearch : Char) : integer;

var i : integer;
    c : char;

begin
  result := 0;
  for i := 1 to UTF8Length(int_Str) do
    begin
      c := int_Str[i];
      if c = CSearch then inc(result);
    end;
end;

{******************************************************************************}

function OnlyDigits(s: string): boolean;

var i : integer;
    c : char;

begin
  result := false;
  for i := 1 to UTF8Length(s) do
    begin
      c := s[i];
      result := (c in ['0'..'9']);
      if not result then break;
    end;
end;

{******************************************************************************}

function  XCharsOnly(s: string; const XChar: TCharSet): string;

var i : integer;
    c : char;

begin
  result := '';
  for i := 1 to UTF8Length(s) do
    begin
      c := s[i];
      if (c in XChar) then result := result + c;
    end;
end;

{******************************************************************************}

function SQLiteDateFormat(d: TDateTime): string;

begin
  result := '"'+FormatDateTime('yyyy-mm-dd', d)+'"';
end;

{******************************************************************************}

function IntToCurrency(i: Longint): string;

begin
  if bTausendertrennung
    then result := Format('%.2n',[i / 100])
    else result := Format('%.2f',[i / 100]);
end;

{******************************************************************************}

function CurrencyToInt(s: string; bEuroM: boolean): Longint;

//Ausgabe in Cent
//Folgende Zahlen werden richtig erkannt
//       0 -->      0 ct             -->      0 ct
//      10 -->     10 ct             -->   1000 ct
//     ,10 -->     10 ct             -->     10 ct
//    0,10 -->     10 ct             -->     10 ct
//    0.10 -->     10 ct             -->     10 ct
//    1,   -->    100 ct             -->    100 ct
//    1.   -->    100 ct             -->    100 ct
//    1.0  -->    100 ct             -->    100 ct
//    1,1  -->    110 ct             -->    110 ct
//    1.11 -->    111 ct             -->    111 ct
//1.000,00 --> 100000 ct             --> 100000 ct
// ein "-" an beliebiger Stelle ergibt eine neg. Zahl

var
  npos,
  CountThousandSeparator,
  CountDecimalSeparator   : integer;
  Euro,
  Cent,
  sHelp                   : String;
  bNeg                    : boolean;

begin
  if s = ''
    then
      begin
        result := 0;
      end
    else
      begin
        try
          sHelp  := s;
          result := 0;

          //Negativ?
          bNeg := (pos('-', sHelp) <> 0);
          if bNeg then sHelp := deletechar(sHelp, '-');

          CountThousandSeparator := CountChar(sHelp, DefaultFormatSettings.ThousandSeparator);
          CountDecimalSeparator  := CountChar(sHelp, DefaultFormatSettings.DecimalSeparator);

          //1.000,00  ---> 1000,00
          if (CountThousandSeparator > 0) and (CountDecimalSeparator > 0) then sHelp := deletechar(sHelp, DefaultFormatSettings.ThousandSeparator);
          //1000.00  ---> 1000,00
          if (CountThousandSeparator = 1) and (CountDecimalSeparator = 0) then sHelp := ReplaceChar(sHelp, DefaultFormatSettings.ThousandSeparator, DefaultFormatSettings.DecimalSeparator);
          //Euromodus: 10 --> 1000
          if (CountThousandSeparator = 0) and (CountDecimalSeparator = 0) and bEuroM then sHelp := sHelp + '00';

          nPos := pos(DefaultFormatSettings.DecimalSeparator, sHelp);
          if nPos = 0
            then //Zahl liegt in Cent vor
              result := strtoint(sHelp)
            else
              begin
                //Damit 21,1 oder 21, richtig umgewandelt wird, werden jetzt erst ein mal 2 0en angefügt
                //falls sie schon vorhanden sind, werden sie später ignoriert
                //für die Zahl ,11 wird noch eine 0 vorne angefügt.
                sHelp := '0' + sHelp + '00';

                //Jetzt sieht die Zahl ggf so aus: '0,1100'
                nPos := pos(DefaultFormatSettings.DecimalSeparator, sHelp); //neu bestimmen
                Euro := copy(sHelp, 1, nPos-1);
                Cent := copy(sHelp, nPos+1, 2);
                result := (strtoint(Euro)*100) + strtoint(Cent);
              end;
        except
          ShowMessage('"'+s+'" hat ein ungültiges Währungsformat');
          result := 0;
        end;

        if bNeg then result := -1 * result;
      end;
end;

{******************************************************************************}

function GetCSVRecordItem(const ItemIndex     : Integer;
                          const CSVRecord     : string;
                          const WordSep       : TCharSet;
                          const WordDelimiter : Char): string;

begin
  result := trim(DeleteChar(ExtractWord(ItemIndex, CSVRecord, WordSep), WordDelimiter));
end;

//************************************************************************************

function WordPosition(const N: Integer; const S: string; const WordSep: TCharSet): Integer;

var
  Count,
  I      : Integer;

begin
  Count  := 1;
  I      := 1;
  if N = 1
    then Result := 1
    else
      begin
        result := 0;
        while (I <= Length(S)) and (Count < N) do
          begin
            while (I <= Length(S)) and not (S[I] in WordSep) do Inc(I);
            if I <= Length(S) then Inc(Count);
            Inc(I); //Auf Zeichen NACH dem Separator gehen
            if Count = N then Result := I;
          end;
      end;
end;

//************************************************************************************

function ExtractWord(const N      : Integer;
                     const S      : string;
                     const WordSep: TCharSet): string;

var
  I,
  Len : Integer;

begin
  Len := 0;
  I   := WordPosition(N, S, WordSep);
  if I <> 0 then
    { find the end of the current word }
    while (I <= Length(S)) and not(S[I] in WordSep) do
      begin
        { add the I'th character to result }
        Inc(Len);
        SetLength(Result, Len);
        Result[Len] := S[I];
        Inc(I);
      end;
  SetLength(Result, Len);
end;

//************************************************************************************

function FieldTypeToString(FT: TFieldtype): string;

//Nur SQLite3 Datentypen sind freigeschaltet

var
  S : string;

begin
  case FT of
    ftString:        S := 'ftString:        String data value (ansistring)';
    ftSmallint:      S := 'ftSmallint:      Small integer value(1 byte, signed)';
    ftInteger:       S := 'ftInteger:       Regular integer value (4 bytes, signed)';
    ftAutoInc:       S := 'ftAutoInc:       Auto-increment integer value (4 bytes)';
    ftLargeint:      S := 'ftLargeint:      Large integer value (8-byte)';
    ftWord:          S := 'ftWord:          Word-sized value(2 bytes, unsigned)';
    ftBoolean:       S := 'ftBoolean:       Boolean value';
    ftFloat:         S := 'ftFloat:         Floating point value (double)';
    ftCurrency:      S := 'ftCurrency:      Currency value (4 decimal points)';
    ftDate:          S := 'ftDate:          Date value';
    ftTime:          S := 'ftTime:          Time value';
    ftDateTime:      S := 'ftDateTime:      Date/Time (timestamp) value';
    ftMemo:          S := 'ftMemo:          Binary text data (no size)';
    ftBytes:         S := 'ftBytes:         Array of bytes value, fixed size (unytped) ';
    ftVarBytes:      S := 'ftVarBytes:      Array of bytes value, variable size (untyped)';
    //ftBCD:           S := 'ftBCD:           Binary Coded Decimal value (DECIMAL and NUMERIC SQL types)';
    //ftBlob:          S := 'ftBlob:          Binary data value (no type, no size)';
    //ftGraphic:       S := 'ftGraphic:       Graphical data value (no size)';
    //ftFmtMemo:       S := 'ftFmtMemo:       Formatted memo ata value (no size)';
    //ftParadoxOle:    S := 'ftParadoxOle:    Paradox OLE field data (no size)';
    //ftDBaseOle:      S := 'ftDBaseOle:      Paradox OLE field data';
    //ftTypedBinary:   S := 'ftTypedBinary:   Binary typed data (no size)';
    //ftCursor:        S := 'ftCursor:        Cursor data value (no size)';
    //ftFixedChar:     S := 'ftFixedChar:     Fixed character array (string)';
    //ftWideString:    S := 'ftWideString:    Widestring (2 bytes per character)';
    //ftADT:           S := 'ftADT:           ADT value';
    //ftArray:         S := 'ftArray:         Array data';
    //ftReference:     S := 'ftReference:     Reference data';
    //ftDataSet:       S := 'ftDataSet:       Dataset data (blob)';
    //ftOraBlob:       S := 'ftOraBlob:       Oracle BLOB data';
    //ftOraClob:       S := 'ftOraClob:       Oracle CLOB data';
    //ftVariant:       S := 'ftVariant:       Variant data value';
    //ftInterface:     S := 'ftInterface:     interface data value';
    //ftIDispatch:     S := 'ftIDispatch:     Dispatch data value';
    //ftGuid:          S := 'ftGuid:          GUID data value';
    //ftTimeStamp:     S := 'ftTimeStamp:     Timestamp data value';
    //ftFMTBcd:        S := 'ftFMTBcd:        Formatted BCD (Binary Coded Decimal) value.';
    //ftFixedWideChar: S := 'ftFixedWideChar: Fixed wide character date (2 bytes per character)';
    //ftWideMemo:      S := 'ftWideMemo:      Widestring memo data';
    else
                     S := 'ftUnknown:       Unknown data type';
  end;
  result := S;
end;

{******************************************************************************}

function myval(int_zahl_str: string) : integer;

{wandelt einen String in eine Integerzahl als Funktion}

var code, int_zahl : integer;

begin
  val(int_zahl_str,int_zahl,code);
  myval := int_zahl;
end;

{******************************************************************************}

Procedure PrintText(sDruckText, sTitel: String; PrtDlg: TPrintDialog);

var
  YPos,
  MaxYPos,
  LineHeight,
  VerticalMargin,
  i,
  LeftMargin   : Integer;
  slText       : TStringlist;

begin
  if PrtDlg.Execute
    then
      begin
        try
          slText := TStringlist.Create;
          slText.Text:=sDruckText;

          Printer.BeginDoc;
          Printer.Title             := sTitel;
          Printer.Canvas.Font.Name  := 'Courier New';
          Printer.Canvas.Font.Size  := 10;
          Printer.Canvas.Font.Color := clBlack;

          LineHeight     := Round(1.2 * Abs(Printer.Canvas.TextHeight('I')));
          LeftMargin     := (printer.PageWidth div 21) * 2; //2cm
          VerticalMargin := 4 * LineHeight;
          YPos           := VerticalMargin;
          MaxYPos        := printer.PageHeight-VerticalMargin;
          for i := 0 to slText.Count-1 do
            begin
              Printer.Canvas.TextOut(LEFTMARGIN, YPos, slText.Strings[i]);
              YPos := YPos + LineHeight;
              if YPos > MaxYPos
                then
                  begin
                    YPos := VerticalMargin;
                    Printer.NewPage;
                  end;
            end;
        finally
          slText.Free;
          Printer.EndDoc;
        end;
      end;
end;

{******************************************************************************}

function ExecSQL(SQLCommand : string; MyQue: TZQuery; ShowUpdateMessage: boolean):integer;

var SQLForm      : TForm;
    MyDS         : TDataSource;
    myGrid       : TDBGrid;
    i,
    w,
    r,
    n            : integer;
    WORKAREA     : TRect;

begin
  try
    screen.cursor  := crhourglass;
    if SQLCommand <> '' // Die Que hat bereits das SQL Statement
      then MyQue.sql.text := SQLCommand
      else SQLCommand     := MyQue.sql.text; //Zurückspeichern

    if      (pos('SELECT', UPPERCASE(SQLCommand)) <> 0) and
       not ((pos('INSERT', UPPERCASE(SQLCommand)) <> 0) or (pos('UPDATE', UPPERCASE(SQLCommand)) <> 0))
      then
        begin
          result              := 0;

          WORKAREA := GETMaxWindowsSize; //Verfügbarer Platz

          //Formular anlegen
          SQLForm             := TForm.Create(application);
          SQLForm.Caption     := 'SQL result';
          SQLForm.Position    := poDesigned;
          SQLForm.Left        := WORKAREA.Left;
          SQLForm.Top         := WORKAREA.Top;
          SQLForm.Width       := WORKAREA.Right;
          SQLForm.Height      := WORKAREA.Bottom;
          SQLForm.BorderStyle := bsSizeAble;
          MyDS                := TDataSource.create(SQLForm);
          myDS.DataSet        := MyQue;
          myGrid              := TDBGrid.create(SQLForm);
          myGrid.Parent       := SQLForm;
          myGrid.Align        := alClient;
          myGrid.DataSource   := myDS;
          myGrid.Options      := myGrid.Options + [dgDisplayMemoText];
          MyQue.Open;

          //Spaltenberenzung
          w := (screen.Width-60) div myGrid.Columns.Count; //Breite für gleichmäßige Aufteilung
          n := 0; //Anzahl der Spalten die kleiner sind
          r := 0; //Restbreite
          for i := 0 to myGrid.Columns.Count-1 do
            begin
              if myGrid.Columns.Items[i].Width < w then
                begin
                  inc(n);
                  r := r + (w- myGrid.Columns.Items[i].Width);
                end;
            end;
          w := w + r div (myGrid.Columns.Count-n); //restlichen Platz verteilen
          for i := 0 to myGrid.Columns.Count-1 do myGrid.Columns.Items[i].Width := min(myGrid.Columns.Items[i].Width, w);

          //Anzeigen
          screen.cursor := crdefault;
          SQLForm.ShowModal;

          //Exportieren?
          if MessageDlg('Sollen die Daten noch exportiert werden?', mtConfirmation, [mbYes, mbNo],0) = mrYes then
            begin
              ExportQueToCSVFile(myQue, sPrintPath+'Export_UTF8.csv', ';', '"',false,  true);
              ExportQueToCSVFile(myQue, sPrintPath+'Export_ANSI.csv', ';', '"', true, false);
            end;

          myQue.Close;
          MyGrid.Free;
          myDS.free;
          SQLForm.Free;
        end
      else
        begin
          //Insert oder
          //Update
          MyQue.ExecSQL;
          result := MyQue.RowsAffected;
          screen.cursor := crdefault;
          if ShowUpdateMessage
            then LogAndShow(format('%d Zeile(n) aktualsiert',[result]))
            else myDebugLN(format('%d Zeile(n) aktualsiert',[result]));
        end;
  except
    on E: Exception
      do
        begin
          screen.cursor := crdefault;
          LogAndShowError('Fehler in SQL-Kommando'#13#13+SQLCommand+#13#13+e.Message);
          result := -1;
        end;
  end;
end;

{******************************************************************************}

procedure GetPrinterMargins();

begin
  with printer do
    Showmessage('PrinterName: '+PrinterName+#13#13+
                'PaperSize.Width: '+RealToStr(PaperSize.Width*25.4/XDPI,0,1)+' mm'+#13+
                'PaperSize.Height: '+RealToStr(PaperSize.Height*25.4/YDPI,0,1)+' mm'+#13#13+
                'PageWidth: '+RealToStr(PageWidth*25.4/XDPI,0,1)+' mm'+#13+
                'PageHeight: '+RealToStr(PageHeight*25.4/YDPI,0,1)+' mm'+#13#13+
                'Linker Rand: '+RealToStr(PaperSize.PaperRect.WorkRect.Left*25.4/XDPI,0,1)+' mm'+#13+
                'Rechter Rand: '+RealToStr((PaperSize.Width-PageWidth-PaperSize.PaperRect.WorkRect.Left)*25.4/XDPI,0,1)+' mm'+#13+
                'Oberer Rand: '+RealToStr(PaperSize.PaperRect.WorkRect.Top*25.4/XDPI,0,1)+' mm'+#13+
                'Unterer Rand: '+RealToStr((PaperSize.Height-PageHeight-PaperSize.PaperRect.WorkRect.Top)*25.4/YDPI,0,1)+' mm'+#13#13+
                'XDPI: '+inttostr(XDPI)+#13+
                'YDPI: '+inttostr(YDPI));
end;

{******************************************************************************}

Function SQL_Where_Add(sWhere, sAdd: string): string;

begin
  result := '';
  if sWhere <> '' then result := sWhere +' and'+#13#10;
  result := result + sAdd;
end;

{******************************************************************************}

Function SQL_Where_Add_OR(sWhere, sAdd: string): string;

begin
  result := '';
  if sWhere <> '' then result := sWhere +' or'+#13#10;
  result := result+sAdd;
end;

{******************************************************************************}

function SQL_Where_Contains(s, Feld: String): string;

begin
  result := format('upper(%s) like ''%%%s%%''',[Feld, uppercase(s)]);
end;

{******************************************************************************}

function SQL_Where_IsNull(Feld: String): string;

begin
  result := '(('+Feld+' is null) or ('+Feld+'=''''))';
end;

{******************************************************************************}

function SQL_QuotedStr(SQL_IN : String): string;

begin
  result := StringReplace(SQL_IN, '"', '""', [rfReplaceAll]);
end;

{******************************************************************************}

function SQL_UTF8UmlautReplace(Feld : String): string;

begin
  result := format('REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(%s,''ö'',''oe''),''Ä'',''Ae''),''Ö'',''Oe''),''ä'',''ae''),''ü'',''ue''),''Ü'',''Ue''),''ß'',''ss''),''é'',''e''),''è'',''e'')', [feld]);
end;

{******************************************************************************}

function SQL_DeleteComment(SQL_IN : String): string;

var
  p1, p2 : integer;

begin
  result := SQL_IN;
  p1 := UTF8Pos('--', result);
  if p1 > 0 then UTF8Delete(result, p1, 9999);

  p1 := UTF8Pos('/*', result);
  p2 := UTF8Pos('*/', result);

  if (p1 > 0) and (p2 > 0) then UTF8Delete(result, p1, p2 - p1 + 2);
  if (p1 > 0) and (p2 = 0) then UTF8Delete(result, p1, 9999);

  result := trim(result);
end;

{******************************************************************************}
  
Procedure myDebugLN(s: string);

begin
  myDebugStart(s);
  myDebugFinish('');
end;

Procedure myDebugStart(s: string);

begin
  sDebug := sDebug + FormatDateTime('dd.mm.yyyy hh:nn:ss.zzz', now()) + ' ' + RemoveLastCRLF(s);
end;

Procedure myDebugAdd(s: string);

begin
  sDebug := sDebug + s;
end;

Procedure myDebugFinish(s: string);

begin
  sDebug := sDebug + RemoveLastCRLF(s)+#13#10;
  if Length(sDebug) > 1000
    then
      begin
        WriteDebug(sDebug);
        sDebug := '';
      end;
end;

Procedure FlushDebug;

begin
  WriteDebug(sDebug);
  sDebug := '';
end;

Procedure WriteDebug(s: string);

var
  NewFileName : String;
  myFile      : TextFile;
  nFileSize   : Int64;
  dtFileDate  : TDateTime;
  
begin
  if bDebug
    then
      begin
        if sDebugFile <> ''
          then
            begin
              //Große Dateien umbenennen
              if fileexists(sDebugFile)
                then
                  begin
                    if GetFileInfo(sDebugFile, dtFileDate, nFileSize)
                      then
                        begin
                          if nFileSize > 1000*1024
                            then
                              begin
                                NewFileName := ExtractFilePath(sDebugFile)+FormatDateTime('yyyymmdd_hhnnss_', now())+ExtractFileName(sDebugFile);
                                renamefile(sDebugFile, NewFileName);
                              end;
                        end;
                  end;

              AssignFile(myFile, UTF8ToSys(sDebugFile));
              try
                if FileExists(UTF8ToSys(sDebugFile))
                  then Append(myFile)
                  else Rewrite(myFile);
                Write(myFile, s);
              finally
                CloseFile(myFile);
              end;
            end;
      end;
end;

{******************************************************************************}

function RealToStr(D:Double; Stellen, Komma:Integer):String;

var s: string;

begin
  Str(D:Stellen:Komma,S);
  RealToStr := S
end;

{******************************************************************************}

function DeleteChar(int_Str:  string; cDel : Char) : string;

begin
  result := StringReplace(int_Str, cDel, '', [rfReplaceAll]);
end;

{******************************************************************************}

function  DeleteChars(int_Str: string; cDel : TCharSet) : string;

var cHelp : Char;

begin
  result := int_Str;
  for cHelp in cDel do result := StringReplace(result, cHelp, '', [rfReplaceAll]);
end;

function  ReplaceChars(int_Str:  string; cDel: TCharSet; cIn : Char) : string;

var cHelp : Char;

begin
  result := int_Str;
  for cHelp in cDel do result := StringReplace(result, cHelp, cIn, [rfReplaceAll]);
end;

{******************************************************************************}

function ReplaceChar(int_Str:  string; cDel, cIn : Char) : string;

begin
  result := StringReplace(int_Str, cDel, cIn, [rfReplaceAll]);
end;

{******************************************************************************}

function year(int_datum: string) : String;
//Erlaubtes Datumsformat 31.12.2001 oder 2001-12-31

var position : integer;

begin
  position := pos('.',int_datum);
  if position <> -1
   then year := copy(int_datum,7,4)
   else year := copy(int_datum,1,4);
end;

//************************************************************************************

function GetProductVersionString: String;

var
  Info: TVersionInfo;
  PV: TFileProductVersion;

begin
  Info := TVersionInfo.Create;
  Info.Load(HINSTANCE);
  PV := Info.FixedInfo.FileVersion;  //Alternativ Info.FixedInfo.ProductVersion
  Result := Format('%d.%d.%d.%d', [PV[0],PV[1],PV[2],PV[3]]);
  Info.Free;
end;

//************************************************************************************

function AgeAtDate(const DateOfBird, DT: TDateTime): Integer;

var
  D1,
  M1,
  Y1,
  D2,
  M2,
  Y2: Word;

begin
  if DT < DateOfBird
    then
      begin
      	Result := -1
      end
    else
      begin
        DecodeDate(DateOfBird, Y1, M1, D1);
        DecodeDate(DT, Y2, M2, D2);
        Result := Y2 - Y1;
        if (M2 < M1) or
          ((M2 = M1) and (D2 < D1))
          then Dec (Result);
      end;
end;

//************************************************************************************

function AgeNow(const DateOfBird: TDateTime): Integer;

begin
  Result := AgeAtDate (DateOfBird, Date);
end;

{******************************************************************************}

function ReadIniVal(FileName, Section, Keyword, DefVal: string; StoreDef : boolean): string;

//Versucht einen Wert zu lesen.
//Wenn er nicht da ist, wird er mit dem DefVal angelegt

Var
 INI    : TINIFile;
 slText : TStringlist;

begin
  INI := TINIFile.Create(UTF8ToSys(FileName));
  slText := TStringlist.Create;
  slText.Delimiter := ',';
  slText.QuoteChar := '"';
  slText.StrictDelimiter := true;
  slText.DelimitedText := INI.ReadString(Section, Keyword, '');
  if slText.Text = ''
    then
      begin
        if StoreDef then INI.WriteString(Section, Keyword, DefVal);
        slText.DelimitedText := DefVal;
      end;
  INI.Free;
  result := RemoveLastCRLF(slText.Text);
  slText.Free;
end;

function  ReadIniBool(FileName, Section, Keyword: string; DefVal, StoreDef : boolean): boolean;

//Versucht einen Wert zu lesen.
//Wenn er nicht da ist, wird er mit dem DefVal angelegt

Var
 INI   : TINIFile;
 sHelp : String;

begin
  INI := TINIFile.Create(UTF8ToSys(FileName));
  sHelp := INI.ReadString(Section, Keyword, '');
  if sHelp = ''
    then
      begin
        if StoreDef then INI.WriteString(Section, Keyword, BoolInString(DefVal));
        Result := DefVal;
      end
    else
      begin
        Result := ('TRUE' = Uppercase(trim(sHelp)));
      end;
  INI.Free;
end;

{==============================================================================}

function  ReadINIInt(FileName, Section, Keyword: String; DefVal: integer): integer;

//Versucht einen Wert zu lesen.
//Wenn er nicht da ist, wird er mit dem DefVal angelegt

Var
 INI   : TINIFile;
 nHelp : integer;

begin
  INI := TINIFile.Create(UTF8ToSys(FileName));
  nHelp := INI.ReadInteger(Section, Keyword, -9999);
  if nHelp = -9999
    then
      begin
        INI.WriteInteger(Section, Keyword, DefVal);
        nHelp := DefVal;
      end;
  INI.Free;
  Result := nHelp;
end;

{==============================================================================}

procedure WriteIniVal(FileName, Section, Keyword, Val: string);

Var
 INI    : TINIFile;
 slText : TStringlist;

begin
  INI := TINIFile.Create(UTF8ToSys(FileName));
  slText := TStringlist.Create;
  slText.Delimiter := ',';
  slText.QuoteChar := '"';
  slText.StrictDelimiter := true;
  slText.Text      := Val;
  INI.WriteString(Section, Keyword, slText.DelimitedText);
  INI.Free;
  slText.Free;
end;

{==============================================================================}

procedure WriteIniBool(FileName, Section, Keyword: string; b: boolean);

Var
 INI: TINIFile;

begin
  INI := TINIFile.Create(UTF8ToSys(FileName));
  INI.WriteString(Section, Keyword, BoolInString(b));
  INI.Free;
end;

procedure WriteIniInt(FileName, Section, Keyword: string; Val: integer);

Var
 INI: TINIFile;

begin
  INI := TINIFile.Create(UTF8ToSys(FileName));
  INI.WriteInteger(Section, Keyword, Val);
  INI.Free;
end;

{==============================================================================}

function AppendChar(Source: String; AddChar: char; len: integer):string;

var
  sHelp : String;
  nHelp : integer;

begin
  sHelp := '';
  if AddChar = ' '
    then
      begin
        nHelp  := len - UTF8Length(Source);
        if nHelp > 0 then sHelp := format('%0:-'+inttostr(nHelp)+'s', [' ']);
        result := Source + sHelp;
      end
    else
      begin
        result := source;
        While not (UTF8Length(result) >= len) do result := result + addchar;      //length gibt bei Sonderzeichen falsche Werte aus
      end;
end;

{******************************************************************************}

function verschluessele(passwort : string):string;

var i    : integer;
    res  : longint;

begin
  res := 0;
  passwort := appendChar(passwort,' ',12);
  for i := 1 to 12 do res := res + ord(passwort[i])*(113-i); //einfach eine sinnlose Formel
  result := inttostr(res);
end;

end.

