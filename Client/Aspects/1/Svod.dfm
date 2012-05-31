object FSvod: TFSvod
  Left = 192
  Top = 103
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1057#1074#1086#1076#1085#1099#1081' '#1086#1090#1095#1077#1090
  ClientHeight = 716
  ClientWidth = 1016
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnKeyUp = FormKeyUp
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 12
    Width = 143
    Height = 13
    Caption = #1057#1074#1086#1076#1085#1099#1081' '#1086#1090#1095#1077#1090' '#1086#1088#1075#1072#1085#1080#1079#1072#1094#1080#1080
  end
  object Label2: TLabel
    Left = 12
    Top = 38
    Width = 162
    Height = 13
    Caption = #1055#1086#1076#1095#1080#1085#1077#1085#1085#1099#1077' '#1089#1074#1086#1076#1085#1099#1077' '#1088#1077#1077#1089#1090#1088#1099
  end
  object DBGrid1: TDBGrid
    Left = 8
    Top = 56
    Width = 1001
    Height = 624
    DataSource = DataSource1
    PopupMenu = PopupMenu1
    ReadOnly = True
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'Number'
        Width = 0
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'date_Time'
        Title.Caption = #1057#1086#1079#1076#1072#1085
        Width = 107
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Patch'
        Title.Caption = #1055#1091#1090#1100' '#1082' '#1092#1072#1081#1083#1091
        Width = 671
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Records'
        Title.Caption = #1040#1089#1087#1077#1082#1090#1086#1074
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Komm'
        Title.Caption = #1057#1090#1072#1090#1091#1089' '#1086#1073#1088#1072#1073#1086#1090#1082#1080
        Width = 117
        Visible = True
      end>
  end
  object Button1: TButton
    Left = 431
    Top = 688
    Width = 154
    Height = 25
    Caption = #1057#1086#1079#1076#1072#1090#1100' '#1089#1074#1086#1076#1085#1099#1081' '#1088#1077#1077#1089#1090#1088
    TabOrder = 1
    OnClick = Button1Click
  end
  object MainPatch: TEdit
    Left = 155
    Top = 8
    Width = 768
    Height = 21
    ReadOnly = True
    TabOrder = 2
    Text = #1047#1072#1076#1072#1081#1090#1077' '#1087#1091#1090#1100' '#1082' '#1092#1072#1081#1083#1091' '#1075#1083#1072#1074#1085#1086#1075#1086' '#1089#1074#1086#1076#1085#1086#1075#1086' '#1088#1077#1077#1089#1090#1088#1072
  end
  object Button2: TButton
    Left = 928
    Top = 8
    Width = 81
    Height = 21
    Caption = #1047#1072#1076#1072#1090#1100' '#1092#1072#1081#1083
    TabOrder = 3
    OnClick = Button2Click
  end
  object TSvod: TADODataSet
    CursorType = ctStatic
    CommandText = 'Select * From MainEcolog ORDER By Number'
    Parameters = <>
    Left = 8
    Top = 688
  end
  object DataSource1: TDataSource
    DataSet = TSvod
    Left = 48
    Top = 688
  end
  object PopupMenu1: TPopupMenu
    Left = 256
    Top = 680
    object N13: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N13Click
    end
    object N14: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N14Click
    end
    object N15: TMenuItem
      Caption = #1042#1074#1077#1088#1093
      OnClick = N15Click
    end
    object N16: TMenuItem
      Caption = #1042#1085#1080#1079
      OnClick = N16Click
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093'|Main_Svod_reestr.dtb'
    Left = 296
    Top = 680
  end
  object Temp: TADOConnection
    LoginPrompt = False
    Left = 360
    Top = 680
  end
  object TempTable: TADODataSet
    Connection = Temp
    Parameters = <>
    Left = 400
    Top = 680
  end
end
