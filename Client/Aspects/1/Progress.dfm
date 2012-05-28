object Prog: TProg
  Left = 192
  Top = 107
  BorderStyle = bsNone
  Caption = 'Prog'
  ClientHeight = 35
  ClientWidth = 371
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 371
    Height = 35
    Align = alClient
    BevelWidth = 2
    TabOrder = 0
    object Label1: TLabel
      Left = 2
      Top = 2
      Width = 367
      Height = 19
      Align = alClient
      Alignment = taCenter
      Caption = #1055#1086#1076#1086#1078#1076#1080#1090#1077' '#1087#1086#1078#1072#1083#1091#1081#1089#1090#1072'...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object PB: TProgressBar
      Left = 2
      Top = 21
      Width = 367
      Height = 12
      Align = alBottom
      Min = 0
      Max = 100
      TabOrder = 0
    end
  end
end
