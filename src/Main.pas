unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XPMan, OleCtrls, SHDocVw, MSHTML, ExtCtrls, IniFiles,
  ComCtrls;

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
    lvUserList: TListView;
    tmrKeepLogin: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure wbBeforeNavigate2(Sender: TObject; const pDisp: IDispatch;
      var URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
      var Cancel: WordBool);
    procedure btnAnalyzeClick(Sender: TObject);
    procedure tmrCountdownTimer(Sender: TObject);
    procedure tmrTicketTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure lvUserListDblClick(Sender: TObject);
    procedure tmrKeepLoginTimer(Sender: TObject);
  private
    { Private declarations }
    FServerTime, FLocalTime, FTicketBeginTime, FTicketEndTime: TDateTime;
    function GetWebDoc: IHTMLDocument2;
    procedure LoadUserList;
    procedure InputUserNameAndPassword(const aUserName, aUserPass: string);
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
  C_LOGIN_URL = 'http://app.wandafilm.com/wandaFilm/login.action';
  C_QUERYCITY_URL = 'http://app.wandafilm.com/wandaFilm/doqueryCitys.action?userId=%s';

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
  vFormatSettings: TFormatSettings;
begin
  // 获取UserID
  vInput := FindInputByName(Self.GetWebDoc, 'userId');
  if (vInput <> nil) and (Trim(vInput.value) <> '') then
    edtUserID.Text := vInput.value;
  // 获取ServerTime
  Self.GetWebDoc.parentWindow.execScript(C_ServerTime_JS, 'javascript');
  vFormatSettings.DateSeparator := '/';
  vFormatSettings.ShortDateFormat := 'yyyy/MM/dd';
  vFormatSettings.TimeSeparator := ':';
  vFormatSettings.ShortTimeFormat := 'HH:mm:ss';
  vInput := FindInputByName(Self.GetWebDoc, C_ServerTimeName);
  if (vInput <> nil) then
    if TryStrToDateTime(vInput.value, vServerTime, vFormatSettings) then
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
    lvUserList.Enabled := False;
    btnAnalyze.Enabled := False;
    tmrCountdown.Enabled := True;
    tmrKeepLogin.Enabled := True;
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
//  FTicketBeginTime := Trunc(Now) + EncodeTime(13, 51, 0, 0);

  Self.LoadUserList;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  wb.Navigate(C_LOGIN_URL);
end;

function TfrmMain.GetWebDoc: IHTMLDocument2;
begin
  Result := wb.Document as IHTMLDocument2;
end;

procedure TfrmMain.InputUserNameAndPassword(const aUserName,
  aUserPass: string);
var
  vUserNameObj, vUserPassObj, vVerifyCodeObj: IHTMLInputElement;
begin
  vUserNameObj := FindInputByName(Self.GetWebDoc, 'userName');
  if vUserNameObj <> nil then
    vUserNameObj.value := aUserName;
  vUserPassObj := FindInputByName(Self.GetWebDoc, 'userPass');
  if vUserPassObj <> nil then
    vUserPassObj.value := aUserPass;
  vVerifyCodeObj := FindInputByName(Self.GetWebDoc, 'verifyCode');
  if vVerifyCodeObj <> nil then
    (vVerifyCodeObj as IHTMLElement2).focus;
end;

procedure TfrmMain.LoadUserList;
const
  C_Section_UserList = 'UserList';
var
  vIniFile: string;
  i: Integer;
  vIni: TIniFile;
  vKeyList: TStrings;
  vValue: string;
  vListItem: TListItem;
begin
  vIniFile := ExtractFilePath(Application.ExeName) + 'WanDa8.ini';
  if not FileExists(vIniFile) then Exit;
  vIni := TIniFile.Create(vIniFile);

  lvUserList.Items.BeginUpdate;
  vKeyList := TStringList.Create;
  try
    vIni.ReadSection(C_Section_UserList, vKeyList);
    for i := 0 to vKeyList.Count - 1 do
    begin
      vValue := vIni.ReadString(C_Section_UserList, vKeyList[i], '');
      vListItem := lvUserList.Items.Add;
      vListItem.Caption := vKeyList[i];
      vListItem.SubItems.Add(vValue);
    end;
  finally
    vKeyList.Free;
    lvUserList.Items.EndUpdate;
  end;
end;

procedure TfrmMain.lvUserListDblClick(Sender: TObject);
var
  vListItem: TListItem;
begin
  vListItem := lvUserList.Selected;
  if vListItem <> nil then
    Self.InputUserNameAndPassword(vListItem.Caption, vListItem.SubItems[0]);
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
    tmrKeepLogin.Enabled := False;
    FTicketEndTime := IncSecond(Now, 10);
    tmrTicket.Enabled := True;
    lblCountdown.Caption := '开始抢票';
  end
  else
    lblCountdown.Caption := FormatDateTime('hh:nn:ss', vCountdown);
end;

procedure TfrmMain.tmrKeepLoginTimer(Sender: TObject);
begin
  wb.Navigate(Format(C_QUERYCITY_URL, [edtUserID.Text]));
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
  if (FLocalTime = 0) and (Pos('login.action', LowerCase(URL)) > 0)
    and (Length(PostData) > 0) then
    FLocalTime := Now;
end;

end.

