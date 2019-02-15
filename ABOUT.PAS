unit ABOUT;

interface

uses Winapi.Windows, System.Classes, System.SysUtils, Vcl.Forms, Vcl.Controls,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Graphics;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    OKButton: TButton;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    procedure FormCreate(Sender: TObject);
  private
    version_: LongRec;
    function GetVersion: string;
  public
    { Public-Deklarationen }
  end;

var
  AboutBox: TAboutBox;

implementation

{$R *.dfm}

procedure TAboutBox.FormCreate(Sender: TObject);
begin
  version_ := LongRec(GetFileVersion(ParamStr(0)));
  version.Caption := 'Version ' + GetVersion;
end;

function TAboutBox.GetVersion: string;
begin
  Result := Format('%d.%d', [version_.Hi, version_.Lo]);
end;

end.
 
