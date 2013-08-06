object frmBaseWindaBrowser: TfrmBaseWindaBrowser
  Left = 920
  Top = 279
  BorderStyle = bsDialog
  Caption = #25250
  ClientHeight = 421
  ClientWidth = 286
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object wb: TWebBrowser
    Left = 0
    Top = 0
    Width = 286
    Height = 421
    Align = alClient
    TabOrder = 0
    OnDocumentComplete = wbDocumentComplete
    ControlData = {
      4C0000008F1D0000832B00000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
end
