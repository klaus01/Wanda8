program WanDa8;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  HtmlParseUtils in 'HtmlParseUtils.pas',
  BaseWindaBrowser in 'BaseWindaBrowser.pas' {frmBaseWindaBrowser};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '�����8Ԫ��ӰƱ';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
