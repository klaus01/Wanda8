unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XPMan, OleCtrls, SHDocVw, MSHTML, ExtCtrls;

type
  TfrmMain = class(TForm)
    wb: TWebBrowser;
    edtUserID: TEdit;
    edtServerTime: TEdit;
    edtLocalTime: TEdit;
    XPManifest1: TXPManifest;
    btnAnalyze: TButton;
    lbl1: TLabel;
    lblCountdown: TLabel;
    tmrCountdown: TTimer;
    tmrTicket: TTimer;
    lbl2: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    lbl3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure wbBeforeNavigate2(Sender: TObject; const pDisp: IDispatch;
      var URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
      var Cancel: WordBool);
    procedure btnAnalyzeClick(Sender: TObject);
    procedure tmrCountdownTimer(Sender: TObject);
    procedure tmrTicketTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    FServerTime, FLocalTime, FTicketBeginTime, FTicketEndTime: TDateTime;
    function GetWebDoc: IHTMLDocument2;
    procedure TicketOnSuccess(Sender: TObject);
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses UrlMon, BaseWindaBrowser, HtmlParseUtils, DateUtils;

const
  C_IPHONE_USERAGENT = 'Mozilla/5.0 (iPhone; CPU iPhone OS 5_0 like Mac OS X) AppleWebKit/534.46 (KHTML, like Gecko) Version/5.1 Mobile/9A334 Safari/7534.48.3';

{-------------------------------------------------------------------------------
  过程名:    SetProcessUserAgent
  作者:      kelei
  日期:      2013.08.03
  参数:      aUserAgent HTTP请求头UserAgent内容
  返回值:    True设置成功
  SetProcessUserAgent('Mozilla/5.0 (iPhone; CPU iPhone OS 5_0 like Mac OS X) AppleWebKit/534.46 (KHTML, like Gecko) Version/5.1 Mobile/9A334 Safari/7534.48.3')
-------------------------------------------------------------------------------}
function SetProcessUserAgent(const aUserAgent: string): Boolean;
begin
  Result := UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, PChar(aUserAgent), Length(aUserAgent), 0) = S_OK;
end;

{$R *.dfm}

{ TfrmMain }

procedure TfrmMain.btnAnalyzeClick(Sender: TObject);
const
  C_ServerTimeName = 'ServerTime';
  C_ServerTime_JS =
'var vElement = document.getElementById("' + C_ServerTimeName + '"); ' +
'if (vElement != null) ' +
'    vElement.parentNode.removeChild(vElement); ' +
'vElement = document.createElement("input"); ' +
'vElement.name = "' + C_ServerTimeName + '"; ' +
'vElement.id = "' + C_ServerTimeName + '"; ' +
'vElement.type = "hidden"; ' +
'vElement.value = datastr; ' +
'document.body.appendChild(vElement); ';
var
  vServerTime: TDateTime;
  vInput: IHTMLInputElement;
  vEndTime: TDateTime;
begin
  // 获取UserID
  vInput := FindInputByName(Self.GetWebDoc, 'userId');
  if (vInput <> nil) and (Trim(vInput.value) <> '') then
    edtUserID.Text := vInput.value;
  // 获取ServerTime
  Self.GetWebDoc.parentWindow.execScript(C_ServerTime_JS, 'javascript');
  vInput := FindInputByName(Self.GetWebDoc, C_ServerTimeName);
  if (vInput <> nil) then
    if TryStrToDateTime(vInput.value, vServerTime) then
    begin
      vEndTime := Now;
      FServerTime := vServerTime - (vEndTime - FLocalTime);
      if FServerTime > FTicketBeginTime then
        FTicketBeginTime := FTicketBeginTime + 1;
      edtServerTime.Text := DateTimeToStr(FServerTime);
      edtLocalTime.Text := DateTimeToStr(FLocalTime);
    end;
  // 开始到倒计时
  if (edtUserID.Text <> '') and (edtServerTime.Text <> '') then
  begin
    btnAnalyze.Enabled := False;
    tmrCountdown.Enabled := True;
  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  Self.Caption := Application.Title;

  SetProcessUserAgent(C_IPHONE_USERAGENT);

  edtUserID.ReadOnly := True;
  edtServerTime.ReadOnly := True;
  edtLocalTime.ReadOnly := True;
  edtUserID.Text := '';
  edtServerTime.Text := '';
  edtLocalTime.Text := '';
  lblCountdown.Caption := '';

  FLocalTime := 0;
  FTicketBeginTime := Trunc(Now) + EncodeTime(9, 59, 58, 0);
//  FTicketBeginTime := Trunc(Now) + EncodeTime(18, 18, 0, 0);
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  wb.Navigate('http://app.wandafilm.com/wandaFilm/login.action');
end;

function TfrmMain.GetWebDoc: IHTMLDocument2;
begin
  Result := wb.Document as IHTMLDocument2;
end;

procedure TfrmMain.TicketOnSuccess(Sender: TObject);
begin
  tmrTicket.Enabled := False;
  lblCountdown.Caption := '抢到了';
end;

procedure TfrmMain.tmrCountdownTimer(Sender: TObject);
var
  vServerTime, vCountdown: TDateTime;
begin
  vServerTime := Now + FServerTime - FLocalTime;
  vCountdown := FTicketBeginTime - vServerTime;
  if vCountdown <= 0 then
  begin
    tmrCountdown.Enabled := False;
    FTicketEndTime := IncSecond(Now, 10);
    tmrTicket.Enabled := True;
    lblCountdown.Caption := '开始抢票';
  end
  else
    lblCountdown.Caption := FormatDateTime('hh:nn:ss', vCountdown);
end;

procedure TfrmMain.tmrTicketTimer(Sender: TObject);
begin
  with TfrmBaseWindaBrowser.Create(Application) do
  begin
    UserID := edtUserID.Text;
    OnSuccess := Self.TicketOnSuccess;
    Show;
  end;
  if Now >= FTicketEndTime then
  begin
    tmrTicket.Enabled := False;
    lblCountdown.Caption := '抢票结束';
  end;
end;

procedure TfrmMain.wbBeforeNavigate2(Sender: TObject;
  const pDisp: IDispatch; var URL, Flags, TargetFrameName, PostData,
  Headers: OleVariant; var Cancel: WordBool);
begin
  if (FLocalTime = 0) and (Pos('login.action', LowerCase(URL)) > 0) then
    FLocalTime := Now;
end;

end.

