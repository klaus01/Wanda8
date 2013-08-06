object frmMain: TfrmMain
  Left = 192
  Top = 130
  BorderStyle = bsDialog
  Caption = 'frmMain'
  ClientHeight = 684
  ClientWidth = 698
  Color = clBtnFace
  Font.Charset = GB2312_CHARSET
  Font.Color = clBlack
  Font.Height = -15
  Font.Name = #24494#36719#38597#40657
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  DesignSize = (
    698
    684)
  PixelsPerInch = 96
  TextHeight = 20
  object lbl1: TLabel
    Left = 442
    Top = 425
    Width = 144
    Height = 20
    Anchors = [akTop, akRight]
    Caption = '3'#12289#33258#21160#25250#31080#20498#35760#26102#65306
  end
  object lblCountdown: TLabel
    Left = 442
    Top = 449
    Width = 103
    Height = 20
    Anchors = [akTop, akRight]
    Caption = 'lblCountdown'
  end
  object lbl2: TLabel
    Left = 442
    Top = 231
    Width = 45
    Height = 20
    Anchors = [akTop, akRight]
    Caption = #29992#25143'ID'
  end
  object Label1: TLabel
    Left = 442
    Top = 287
    Width = 75
    Height = 20
    Anchors = [akTop, akRight]
    Caption = #26381#21153#22120#26102#38388
  end
  object Label2: TLabel
    Left = 442
    Top = 343
    Width = 60
    Height = 20
    Anchors = [akTop, akRight]
    Caption = #26412#22320#26102#38388
  end
  object lbl3: TLabel
    Left = 442
    Top = 23
    Width = 214
    Height = 20
    Anchors = [akTop, akRight]
    Caption = '1'#12289#20808#30331#24405'('#21452#20987#21015#34920#21487#33258#21160#36755#20837')'
  end
  object wb: TWebBrowser
    Left = 0
    Top = 0
    Width = 409
    Height = 684
    Align = alLeft
    TabOrder = 0
    OnBeforeNavigate2 = wbBeforeNavigate2
    ControlData = {
      4C000000452A0000B24600000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
  object edtUserID: TEdit
    Left = 442
    Top = 255
    Width = 225
    Height = 28
    Anchors = [akTop, akRight]
    ImeName = 'Chinese (Simplified) - US Keyboard'
    TabOrder = 2
    Text = 'edtUserID'
  end
  object edtServerTime: TEdit
    Left = 442
    Top = 311
    Width = 225
    Height = 28
    Anchors = [akTop, akRight]
    ImeName = 'Chinese (Simplified) - US Keyboard'
    TabOrder = 3
    Text = 'edtServerTime'
  end
  object edtLocalTime: TEdit
    Left = 442
    Top = 367
    Width = 225
    Height = 28
    Anchors = [akTop, akRight]
    ImeName = 'Chinese (Simplified) - US Keyboard'
    TabOrder = 4
    Text = 'edtLocalTime'
  end
  object btnAnalyze: TButton
    Left = 442
    Top = 200
    Width = 225
    Height = 25
    Anchors = [akTop, akRight]
    Caption = '2'#12289#24050#25104#21151#30331#24405#65292#28857#20987#33719#21462#20449#24687
    TabOrder = 1
    OnClick = btnAnalyzeClick
  end
  object lvUserList: TListView
    Left = 442
    Top = 48
    Width = 225
    Height = 113
    Anchors = [akTop, akRight]
    Columns = <
      item
        AutoSize = True
        Caption = #29992#25143#21517
      end
      item
        AutoSize = True
        Caption = #23494#30721
      end>
    RowSelect = True
    TabOrder = 5
    ViewStyle = vsReport
    OnDblClick = lvUserListDblClick
  end
  object XPManifest1: TXPManifest
    Left = 432
    Top = 16
  end
  object tmrCountdown: TTimer
    Enabled = False
    OnTimer = tmrCountdownTimer
    Left = 488
    Top = 16
  end
  object tmrTicket: TTimer
    Enabled = False
    Interval = 200
    OnTimer = tmrTicketTimer
    Left = 552
    Top = 16
  end
  object tmrKeepLogin: TTimer
    Enabled = False
    Interval = 30000
    OnTimer = tmrKeepLoginTimer
    Left = 608
    Top = 16
  end
end
