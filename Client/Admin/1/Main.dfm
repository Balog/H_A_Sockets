object Form1: TForm1
  Left = 192
  Top = 107
  Width = 427
  Height = 309
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 5
    Top = 28
    Width = 39
    Height = 13
    Caption = #1051#1086#1075#1080#1085#1099
  end
  object Label2: TLabel
    Left = 240
    Top = 30
    Width = 80
    Height = 13
    Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
  end
  object Label3: TLabel
    Left = 5
    Top = 7
    Width = 25
    Height = 13
    Caption = #1056#1086#1083#1100
  end
  object Label4: TLabel
    Left = 243
    Top = 8
    Width = 65
    Height = 13
    Caption = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093
  end
  object Role: TComboBox
    Left = 40
    Top = 3
    Width = 196
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 0
  end
  object Users: TListBox
    Left = 2
    Top = 45
    Width = 233
    Height = 214
    ItemHeight = 13
    PopupMenu = PopupMenu1
    TabOrder = 1
  end
  object Otdels: TListBox
    Left = 239
    Top = 45
    Width = 177
    Height = 214
    ItemHeight = 13
    PopupMenu = FreeOtdel
    TabOrder = 2
  end
  object CBDatabase: TComboBox
    Left = 315
    Top = 2
    Width = 101
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 3
  end
  object Database1: TADOConnection
    Left = 24
    Top = 64
  end
  object PopupMenu1: TPopupMenu
    Left = 56
    Top = 64
    object N1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103
    end
    object N2: TMenuItem
      Caption = #1048#1079#1084#1077#1085#1080#1090#1100' '#1083#1086#1075#1080#1085
    end
    object N3: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103
    end
    object N4: TMenuItem
      Caption = #1048#1079#1084#1077#1085#1080#1090#1100' '#1087#1072#1088#1086#1083#1100
    end
  end
  object FreeOtdel: TPopupMenu
    Left = 88
    Top = 64
  end
  object MainMenu1: TMainMenu
    Left = 328
    Top = 16
    object N5: TMenuItem
      Caption = #1063#1090#1077#1085#1080#1077
    end
    object N6: TMenuItem
      Caption = #1047#1072#1087#1080#1089#1100
    end
    object N7: TMenuItem
      Caption = #1046#1091#1088#1085#1072#1083' '#1089#1086#1073#1099#1090#1080#1081
    end
    object N10: TMenuItem
      Caption = #1055#1086#1084#1086#1097#1100
    end
    object N8: TMenuItem
      Caption = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
    end
    object N9: TMenuItem
      Caption = #1042#1099#1093#1086#1076
    end
  end
  object DBTimer: TTimer
    Enabled = False
    Interval = 100
    Left = 128
    Top = 64
  end
  object ClientSocket: TClientSocket
    Active = False
    ClientType = ctNonBlocking
    Port = 0
    OnRead = ClientSocketRead
    OnWrite = ClientSocketWrite
    Left = 32
    Top = 128
  end
  object ActionManager1: TActionManager
    Left = 72
    Top = 128
    object RegForm: TAction
      Caption = 'RegForm'
      OnExecute = RegFormExecute
    end
  end
end
