object FProg: TFProg
  Left = 203
  Top = 162
  BorderStyle = bsNone
  Caption = 'FProg'
  ClientHeight = 30
  ClientWidth = 319
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnHide = FormHide
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Progress: TProgressBar
    Left = 0
    Top = 17
    Width = 319
    Height = 13
    Align = alBottom
    Min = 0
    Max = 9
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 319
    Height = 16
    Align = alTop
    BevelOuter = bvLowered
    TabOrder = 1
    object Label1: TLabel
      Left = 1
      Top = 1
      Width = 317
      Height = 14
      Align = alClient
    end
  end
  object Timer1: TTimer
    Interval = 10000
    OnTimer = Timer1Timer
    Left = 72
  end
end
