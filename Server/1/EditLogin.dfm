object EditLogins: TEditLogins
  Left = 657
  Top = 176
  BorderStyle = bsDialog
  Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1077
  ClientHeight = 139
  ClientWidth = 224
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnKeyUp = FormKeyUp
  OnShow = FormShow
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
    OnChange = Pass1Change
  end
  object Pass2: TMaskEdit
    Left = 48
    Top = 72
    Width = 169
    Height = 21
    PasswordChar = '*'
    TabOrder = 2
    Text = 'MaskEdit1'
    OnChange = Pass1Change
  end
  object Button1: TButton
    Left = 74
    Top = 104
    Width = 75
    Height = 25
    Caption = #1047#1072#1087#1080#1089#1072#1090#1100
    Enabled = False
    TabOrder = 3
    OnClick = Button1Click
  end
end