object Pass: TPass
  Left = 213
  Top = 182
  BorderStyle = bsNone
  Caption = 'Pass'
  ClientHeight = 54
  ClientWidth = 242
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 0
    Width = 242
    Height = 54
    Align = alClient
    Shape = bsFrame
    Style = bsRaised
  end
  object Label1: TLabel
    Left = 6
    Top = 10
    Width = 31
    Height = 13
    Caption = #1051#1086#1075#1080#1085
  end
  object Label2: TLabel
    Left = 6
    Top = 33
    Width = 38
    Height = 13
    Caption = #1055#1072#1088#1086#1083#1100
  end
  object CbLogin: TComboBox
    Left = 47
    Top = 5
    Width = 159
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 0
  end
  object EdPass: TEdit
    Left = 47
    Top = 29
    Width = 159
    Height = 21
    PasswordChar = '*'
    TabOrder = 1
    OnKeyPress = EdPassKeyPress
  end
  object Button1: TButton
    Left = 210
    Top = 6
    Width = 27
    Height = 43
    Caption = 'OK'
    TabOrder = 2
    OnClick = Button1Click
  end
end
