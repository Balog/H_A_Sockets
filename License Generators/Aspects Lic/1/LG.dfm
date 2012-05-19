object Form1: TForm1
  Left = 192
  Top = 107
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #1043#1077#1085#1077#1088#1072#1090#1086#1088' '#1083#1080#1094#1077#1085#1079#1080#1081
  ClientHeight = 107
  ClientWidth = 224
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
    Left = 8
    Top = 55
    Width = 139
    Height = 13
    Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1077#1081
  end
  object Label2: TLabel
    Left = 47
    Top = 1
    Width = 124
    Height = 29
    Caption = #1040#1057#1055#1045#1050#1058#1067
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Edit1: TEdit
    Left = 152
    Top = 52
    Width = 49
    Height = 21
    ReadOnly = True
    TabOrder = 0
    Text = '0'
  end
  object UpDown1: TUpDown
    Left = 201
    Top = 52
    Width = 15
    Height = 21
    Associate = Edit1
    Min = 0
    Max = 99
    Position = 0
    TabOrder = 1
    Wrap = False
  end
  object CheckBox1: TCheckBox
    Left = 8
    Top = 29
    Width = 209
    Height = 17
    Caption = #1054#1075#1088#1072#1085#1080#1095#1077#1085#1072#1103' '#1083#1080#1094#1077#1085#1079#1080#1103
    Checked = True
    State = cbChecked
    TabOrder = 2
    OnClick = CheckBox1Click
  end
  object Button1: TButton
    Left = 52
    Top = 77
    Width = 121
    Height = 25
    Caption = #1057#1086#1079#1076#1072#1090#1100' '#1083#1080#1094#1077#1085#1079#1080#1102
    TabOrder = 3
    OnClick = Button1Click
  end
end
