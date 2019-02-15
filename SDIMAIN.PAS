unit SDIMAIN;

interface

uses Winapi.Windows, System.Classes, Vcl.Graphics, Vcl.Forms, Vcl.Controls,
  Vcl.Menus, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.ImgList, Vcl.StdActns, Vcl.ActnList, Vcl.ToolWin, System.AnsiStrings,
  Data.DB, Data.Win.ADODB, system.SysUtils;

type

  TSDIAppForm = class(TForm)
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    ToolBar1: TToolBar;
    ToolButton9: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ActionList1: TActionList;
    FileNew1: TAction;
    FileOpen1: TAction;
    FileSave1: TAction;
    FileSaveAs1: TAction;
    FileExit1: TAction;
    EditCut1: TEditCut;
    EditCopy1: TEditCopy;
    EditPaste1: TEditPaste;
    HelpAbout1: TAction;
    StatusBar: TStatusBar;
    ImageList1: TImageList;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    FileNewItem: TMenuItem;
    FileOpenItem: TMenuItem;
    FileSaveItem: TMenuItem;
    FileSaveAsItem: TMenuItem;
    N1: TMenuItem;
    FileExitItem: TMenuItem;
    Edit1: TMenuItem;
    CutItem: TMenuItem;
    CopyItem: TMenuItem;
    PasteItem: TMenuItem;
    Help1: TMenuItem;
    HelpAboutItem: TMenuItem;
    ComboBox1: TComboBox;
    Label1: TLabel;
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    procedure FileNew1Execute(Sender: TObject);
    procedure FileOpen1Execute(Sender: TObject);
    procedure FileSave1Execute(Sender: TObject);
    procedure FileExit1Execute(Sender: TObject);
    procedure HelpAbout1Execute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    ADODataset: TADODataset;
    Connection: TADOConnection;
    DataSource: TDataSource;
    FProvider: string;
    FTableList: Tstringlist;
    FFieldList: TStringList;
    function Connect(const Path: string; const LoginPrompt: boolean = False): boolean;
    procedure getTables(const list: TStrings);
    procedure RemoveSpace(var str: String; StartUpperCase: boolean = true);
    function embrace(const s: String) : string;
    function FieldType (ft: TFieldType): string;
    function FieldZero (ft: TFieldType): string;
  public

  end;

var
  SDIAppForm: TSDIAppForm;

implementation
uses About;

{$R *.dfm}

var     Names: TStringList;

procedure TSDIAppForm.FileNew1Execute(Sender: TObject);
begin
  { Nichts ausführen }
end;

procedure TSDIAppForm.FileOpen1Execute(Sender: TObject);
begin
  if OpenDialog.Execute then
    if not Connect(openDialog.FileName) then
      if not Connect(Opendialog.FileName) then
        showmessage('Could not open ' + Opendialog.FileName);
  getTables(ComboBox1.Items);
  ComboBox1.ItemIndex := 0;
end;

procedure TSDIAppForm.FileSave1Execute(Sender: TObject);
begin
  SaveDialog.Execute;
end;

procedure TSDIAppForm.FormCreate(Sender: TObject);
var s: string;

begin
  Names := TStringList.Create;
  GetProviderNames(Names);
  Names.Sort;
//  showmessage(inttostr(Names.IndexOf('Microsoft.ACE.OLEDB.12.0')));
  for s in Names do
    if (pos('Microsoft.ACE.OLEDB.12.0', s)) <> 0 then
      FProvider := s;
  if FProvider = '' then
    for s in Names do
      if (pos('Microsoft.ACE.OLEDB', s)) <> 0 then
        FProvider := s;
  FTableList := TStringlist.Create;
  FFieldList := TStringlist.Create;
  FFieldList.Append('');
  FFieldList.Sorted := False;
  ADODataSet := TADODataset.Create(Application);
end;

procedure TSDIAppForm.FormResize(Sender: TObject);
begin
 Memo1.Width := ClientWidth;
 Memo1.Height := StatusBar.Top - Memo1.Top;
end;

procedure TSDIAppForm.FileExit1Execute(Sender: TObject);
begin
  Close;
end;

procedure TSDIAppForm.HelpAbout1Execute(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TSDIAppForm.Button1Click(Sender: TObject);
var s, s1, tableName, sqlstr: string;
    i, c: integer;
    sa: TStringList;
begin
  if not assigned(connection) then exit;

  sa := TStringList.Create;
  FFieldList.Clear;
  connection.GetFieldNames(s, FFieldList);
  memo1.Clear;
  s := Trim(ComboBox1.Text);
  tableName := embrace(s);
  connection.GetFieldNames(s, FFieldList);
  ADODataset.Close;
  ADODataset.Connection := connection;
  ADODataset.CommandText := 'SELECT * FROM ' + tableName + ';';
  ADODataset.Open;
  RemoveSpace(s);
  memo1.Lines.Append('T' + s + ' = class');
  memo1.Lines.Append('protected');
  memo1.Lines.Append('const');
  s1 := embrace(AdoDataset.FieldDefs[0].Name);
  c := 1;
  while c < FFieldList.Count do
  begin
    s1 := s1 + ', ' + embrace(AdoDataset.FieldDefs[c].Name);
    inc(c);
  end;
  c := 0;
  memo1.Lines.Append('  SQLstr = ''SELECT ' + s1 + ' FROM ' + tableName + ''';');
  memo1.Lines.Append('var');
  while c < FFieldList.Count do
  begin
    s := AdoDataset.FieldDefs[c].Name;
    RemoveSpace(s);
    sa.Append(s);
    s1 := FieldType(AdoDataset.FieldDefs[c].DataType);
    memo1.Lines.Append('  F' + s + ': ' + s1 + ';');
    inc(c);
  end;
  c := 0;
  while c < FFieldList.Count do
  begin
    s1 := FieldType(AdoDataset.FieldDefs[c].DataType);
    memo1.Lines.Append('  procedure Set' + sa.Strings[c] + '(const Value: ' + s1 + ');');
    inc(c);
  end;
  c := 0;
  memo1.Lines.Append('public');
  while c < FFieldList.Count do
  begin
    s1 := FieldType(AdoDataset.FieldDefs[c].DataType);
    memo1.Lines.Append('  property ' + sa.Strings[c] + ': ' + s1 + ' read F' +
                        sa.Strings[c] + ' write Set' + sa.Strings[c] + ';');
    inc(c);
  end;
  memo1.Lines.Append('  procedure Clear;');
  memo1.Lines.Append('');
  memo1.Lines.Append('end;');


end;

procedure TSDIAppForm.Button2Click(Sender: TObject);
var s, s1, s2, tableName: string;
    i, c: integer;
    sa: TStringList;
begin
  if not assigned(connection) then exit;
  sa := TStringList.Create;
  FFieldList.Clear;
  connection.GetFieldNames(s, FFieldList);
  memo1.Clear;
  s := Trim(ComboBox1.Text);
  tableName := embrace(s);
  connection.GetFieldNames(s, FFieldList);
  ADODataset.Close;
  ADODataset.Connection := connection;
  ADODataset.CommandText := 'SELECT * FROM ' + tableName + ';';
  ADODataset.Open;
  RemoveSpace(s);
  memo1.Lines.Append('{ T' + s + ' }');
  s1 := embrace(AdoDataset.FieldDefs[0].Name);
  c := 1;
  while c < FFieldList.Count do
  begin
    s1 := s1 + ', ' + embrace(AdoDataset.FieldDefs[c].Name);
    inc(c);
  end;
  c := 0;
  while c < FFieldList.Count do
  begin
    s2 := AdoDataset.FieldDefs[c].Name;
    RemoveSpace(s2);
    sa.Append(s2);
    s1 := FieldType(AdoDataset.FieldDefs[c].DataType);
//    memo1.Lines.Append('  F' + s + ': ' + s1 + ';');
    inc(c);
  end;
  c := 0;
  while c < FFieldList.Count do
  begin
    s1 := FieldType(AdoDataset.FieldDefs[c].DataType);
    memo1.Lines.Append('');
    memo1.Lines.Append('procedure T' + s + '.Set' + sa.Strings[c] + '(const Value: ' + s1 + ');');
    memo1.Lines.Append('begin');
    memo1.Lines.Append('  F' + sa.Strings[c] + ' := Value;');
    memo1.Lines.Append('end;');
    inc(c);
  end;

  c := 0;
  memo1.Lines.Append('');
  memo1.Lines.Append('procedure T' + s + '.Clear;');
  memo1.Lines.Append('begin');
  while c < FFieldList.Count do
  begin
    s1 := FieldZero(AdoDataset.FieldDefs[c].DataType);
    memo1.Lines.Append('  F' + sa.Strings[c] + s1);
    inc(c);
  end;
  memo1.Lines.Append('end;');

end;

function TSDIAppForm.Connect(const Path: string; const LoginPrompt: boolean = False): boolean;
  var constr, a, b: string;
begin
  result := true;
  a := 'Provider=' + FProvider + ';User ID=Admin;Data Source=';
  b := ';Mode=Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";' +
    'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;' +
    'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";' +
    'Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;' +
    'Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;' +
    'Jet OLEDB:Bypass UserInfo Validation=False;';
  constr := a + Path + b;
  if not assigned(Connection) then
    Connection := TADOConnection.Create(Application);
  Connection.Close;
  if FProvider = '' then
    showmessage('Missing Provider')
  else
    if not Connection.Connected then
    begin
      Connection.ConnectionString :=  constr;
      Connection.ConnectOptions := coConnectUnspecified;
      Connection.LoginPrompt := LoginPrompt;
      Connection.Open;
    end;
  result := Connection.Connected;
end;

procedure TSDIAppForm.getTables(const list: TStrings);
begin
  list.Clear;
  Connection.GetTableNames(list, false);
end;

procedure TSDIAppForm.RemoveSpace(var str: String; StartUpperCase: boolean = true);
var s, s1: string;
    i: integer;
begin
  str := trim(str);
  if length(str) = 0 then
    str := ''
  else
  begin
    if StartUpperCase then
      s := uppercase(copy(str, 1, 1)) + copy(str, 2, Length(str) - 1)
    else
      s := str;
    repeat
      i := pos(' ', s);
      if i <>0 then
      begin
        s1 := copy(s, i+1, 1);
        if StartUpperCase then
        begin
          s1 := uppercase(s1);
          delete(s, i, 2);
          insert(s1, s, i);
        end
        else
          delete(s, i, 2);
      end;
    until i = 0;
    str := s;
  end;
end;

function TSDIAppForm.embrace(const s: String) : string;
begin
  if pos(' ', s) <> 0 then
    result := '[' + s  + ']'
  else
    result := s;
end;

function TSDIAppForm.FieldType (ft: TFieldType): string;
begin
  case ft of
    ftString:   result := 'string';
    ftInteger:  result := 'integer';
    ftSmallInt: result := 'smallInt';
    ftWord:     result := 'word';
    ftBoolean:  result := 'boolean';
    ftFloat:    result := 'double';
    ftWideString: result := 'string';
    ftCurrency: result := 'currency';
    ftDate:     result := 'TDate';
    ftDateTime: result := 'TDate';
    ftAutoinc:  result := 'integer';
    ftBCD:      result := 'TCurrency';
    ftWideMemo: result := 'TStringList';
    else
    begin
      showmessage('Not implemented: ' + inttostr(integer(ft)));
      result := '???';
    end;
  end;
end;

function TSDIAppForm.FieldZero (ft: TFieldType): string;
const isZero = ' : = 0;';
begin
  case ft of
    ftString:   result := ' := '''';';
    ftSmallInt: result := isZero;
    ftInteger:  result := isZero;
    ftWord:     result := isZero;
    ftBoolean:  result := ' := false;';
    ftFloat:    result := isZero;
    ftWideString: result := ' := '''';';
    ftCurrency: result := isZero;
    ftDate:     result := isZero;
    ftDateTime: result := isZero;
    ftAutoinc:  result := isZero;
    ftBCD:      result := isZero;
    ftWideMemo: result := '.Clear;';
    else
    begin
      showmessage('Not implemented: ' + inttostr(integer(ft)));
      result := '???';
    end;
  end;
end;

{TableInfo}

end.
