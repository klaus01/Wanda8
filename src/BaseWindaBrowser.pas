unit BaseWindaBrowser;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw, MSHTML;

type
  TfrmBaseWindaBrowser = class(TForm)
    wb: TWebBrowser;
    procedure FormShow(Sender: TObject);
    procedure wbDocumentComplete(Sender: TObject; const pDisp: IDispatch;
      var URL: OleVariant);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FOnSuccess: TNotifyEvent;
    { Private declarations }
    function GetWebDoc: IHTMLDocument2;
  public
    UserID: string;
    { Public declarations }
    property OnSuccess: TNotifyEvent read FOnSuccess write FOnSuccess;
  end;

implementation

const
  C_CITYCODE = '4905168908';
  C_QUERTYCITY_URL = 'http://app.wandafilm.com/wandaFilm/doqueryCitys.action?userId=%s';
  C_TICKET_URL = 'http://app.wandafilm.com/wandaFilm/ticket.action?cityCode=%s%s&userId=%s';

{$R *.dfm}

function GetRandomNumberStr(const aLength: Byte): string;
var
  i: Integer;
begin
  Randomize;
  Result := '';
  for i := 1 to aLength do
    Result := Result + IntToStr(Random(9));
end;

{ TfrmBaseWindaBrowser }

procedure TfrmBaseWindaBrowser.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmBaseWindaBrowser.FormShow(Sender: TObject);
var
  vFlags, vHeaders, vTargetFrameName, vPostData: OLEVariant;
  vRef: string;
  vUrl, vRandomStr: string;
begin
  vRandomStr := GetRandomNumberStr(51);
  vUrl := Format(C_TICKET_URL, [vRandomStr, C_CITYCODE, UserID]);
  vFlags := '1';
  vTargetFrameName := '';
  vPostData := '';
  vRef := Format(C_QUERTYCITY_URL, [UserID]);
  vHeaders := 'Referer: ' + vRef + #13#10;
  wb.Navigate(vUrl, vFlags, vTargetFrameName, vPostData, vHeaders);
end;

function TfrmBaseWindaBrowser.GetWebDoc: IHTMLDocument2;
begin
  Result := wb.Document as IHTMLDocument2;
end;

procedure TfrmBaseWindaBrowser.wbDocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  if Pos('ÇÀµ½ÁË', Self.GetWebDoc.body.innerText) > 0 then
  begin
    if Assigned(FOnSuccess) then
      FOnSuccess(Self);
  end
  else
    Self.Close;
end;

end.
