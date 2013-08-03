program WanDa8;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  HtmlParseUtils in 'HtmlParseUtils.pas',
  BaseWindaBrowser in 'BaseWindaBrowser.pas' {frmBaseWindaBrowser};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '抢万达8元电影票';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
