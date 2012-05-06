object EditLogins: TEditLogins
  Left = 192
  Top = 107
  BorderStyle = bsDialog
  Caption = 'EditLogins'
  ClientHeight = 135
  ClientWidth = 222
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 6
    Top = 14
    Width = 31
    Height = 13
    Caption = #1051#1086#1075#1080#1085
  end
  object Label2: TLabel
    Left = 6
    Top = 54
    Width = 38
    Height = 13
    Caption = #1055#1072#1088#1086#1083#1100
  end
  object Log: TEdit
    Left = 48
    Top = 8
    Width = 169
    Height = 21
    TabOrder = 0
    Text = 'Log'
  end
  object Pass1: TMaskEdit
    Left = 48
    Top = 48
    Width = 169
    Height = 21
    PasswordChar = '*'
    TabOrder = 1
    Text = 'Pass1'
  end
  object Pass2: TMaskEdit
    Left = 48
    Top = 72
    Width = 169
    Height = 21
    PasswordChar = '*'
    TabOrder = 2
    Text = 'MaskEdit1'
  end
  object Button1: TButton
    Left = 74
    Top = 104
    Width = 75
    Height = 25
    Caption = #1047#1072#1087#1080#1089#1072#1090#1100
    Enabled = False
    TabOrder = 3
  end
end
