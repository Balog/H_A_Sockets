object Prog: TProg
  Left = 192
  Top = 107
  BorderStyle = bsNone
  Caption = 'Prog'
  ClientHeight = 32
  ClientWidth = 371
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
    Left = 75
    Top = 0
    Width = 220
    Height = 20
    Caption = #1055#1086#1076#1086#1078#1076#1080#1090#1077' '#1087#1086#1078#1072#1083#1091#1081#1089#1090#1072'...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object PB: TProgressBar
    Left = 0
    Top = 20
    Width = 371
    Height = 12
    Align = alBottom
    Min = 0
    Max = 100
    TabOrder = 0
  end
end
