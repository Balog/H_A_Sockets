object Form1: TForm1
  Left = 192
  Top = 107
  BorderStyle = bsDialog
  Caption = 
    #1055#1088#1086#1075#1088#1072#1084#1084#1072' '#1082#1086#1085#1074#1077#1088#1090#1072#1094#1080#1080' '#1073#1072#1079' '#1076#1072#1085#1085#1099#1093' '#1101#1082#1086#1083#1086#1075#1080#1095#1077#1089#1082#1080#1093' '#1072#1089#1087#1077#1082#1090#1086#1074' '#1074#1077#1088#1089#1080#1080' 2' +
    '011 '#1075' '#1082' '#1074#1077#1088#1089#1080#1080' 2012 '#1075
  ClientHeight = 300
  ClientWidth = 862
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
  object Label1: TLabel
    Left = 9
    Top = 6
    Width = 237
    Height = 13
    Caption = #1060#1072#1081#1083#1099' Eco_Aspects '#1086#1090' '#1088#1072#1079#1085#1099#1093' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1077#1081
  end
  object Label2: TLabel
    Left = 8
    Top = 224
    Width = 213
    Height = 13
    Caption = #1060#1072#1081#1083' Reference '#1086#1090' '#1075#1083#1072#1074#1085#1086#1075#1086' '#1089#1087#1077#1094#1080#1072#1083#1080#1089#1090#1072
  end
  object AspRefs: TEdit
    Left = 8
    Top = 240
    Width = 673
    Height = 21
    ReadOnly = True
    TabOrder = 0
  end
  object Button1: TButton
    Left = 792
    Top = 238
    Width = 67
    Height = 25
    Caption = #1059#1082#1072#1079#1072#1090#1100
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 382
    Top = 272
    Width = 97
    Height = 25
    Caption = #1050#1086#1085#1074#1077#1088#1090#1080#1088#1086#1074#1072#1090#1100
    TabOrder = 2
    OnClick = Button2Click
  end
  object AspPatch: TDBGrid
    Left = 8
    Top = 24
    Width = 849
    Height = 193
    DataSource = DataSource1
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
    PopupMenu = PopupMenu1
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'Patch'
        Title.Caption = #1055#1091#1090#1100
        Width = 710
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Result'
        Title.Caption = #1057#1090#1072#1090#1091#1089
        Width = 105
        Visible = True
      end>
  end
  object Edit1: TEdit
    Left = 688
    Top = 240
    Width = 97
    Height = 21
    ReadOnly = True
    TabOrder = 4
  end
  object PopupMenu1: TPopupMenu
    Left = 16
    Top = 72
    object N1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N2Click
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #1040#1089#1087#1077#1082#1090#1099'|*.dtb'
    Left = 56
    Top = 72
  end
  object OpenDialog2: TOpenDialog
    Filter = #1041#1072#1079#1072' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1086#1074'|*.dtb'
    Left = 272
    Top = 248
  end
  object Database: TADOConnection
    Left = 288
    Top = 8
  end
  object Table: TADODataSet
    Connection = Database
    Parameters = <>
    Left = 328
    Top = 8
  end
  object DataSource1: TDataSource
    DataSet = Table
    Left = 368
    Top = 8
  end
  object TempFrom: TADOConnection
    Left = 576
    Top = 8
  end
  object TempTo: TADOConnection
    Left = 616
    Top = 8
  end
  object AspectTo: TADOConnection
    Left = 688
    Top = 8
  end
end
