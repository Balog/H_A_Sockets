object SZn: TSZn
  Left = 235
  Top = 283
  BorderStyle = bsDialog
  Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1082#1088#1080#1090#1077#1088#1080#1077#1074
  ClientHeight = 355
  ClientWidth = 1002
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PopupMenu = PopupMenu1
  Position = poDesktopCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DBGrid1: TDBGrid
    Left = 1
    Top = 4
    Width = 998
    Height = 317
    DataSource = DataSource1
    Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs]
    PopupMenu = PopupMenu1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1079#1085#1072#1095#1080#1084#1086#1089#1090#1080
        Title.Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = #1052#1080#1085' '#1075#1088#1072#1085#1080#1094#1072
        Visible = True
      end
      item
        Expanded = False
        FieldName = #1050#1088#1080#1090#1077#1088#1080#1081'1'
        PickList.Strings = (
          #1044#1072
          #1053#1077#1090)
        Title.Caption = #1047#1085#1072#1095#1080#1084#1086#1089#1090#1100
        Width = 69
        Visible = True
      end
      item
        Expanded = False
        FieldName = #1053#1077#1086#1073#1093#1086#1076#1080#1084#1072#1103' '#1084#1077#1088#1072
        Width = 725
        Visible = True
      end>
  end
  object Button1: TButton
    Left = 464
    Top = 327
    Width = 75
    Height = 25
    Caption = #1047#1072#1082#1088#1099#1090#1100
    TabOrder = 1
    OnClick = Button1Click
  end
  object DataSource1: TDataSource
    DataSet = ADODataSet1
    OnUpdateData = DataSource1UpdateData
    Left = 8
    Top = 327
  end
  object ADODataSet1: TADODataSet
    CursorType = ctStatic
    BeforeInsert = ADODataSet1BeforeInsert
    AfterInsert = ADODataSet1AfterInsert
    BeforeEdit = ADODataSet1BeforeEdit
    BeforePost = ADODataSet1BeforePost
    AfterPost = ADODataSet1AfterPost
    CommandText = 'select * from '#1047#1085#1072#1095#1080#1084#1086#1089#1090#1100' Order By ['#1052#1072#1082#1089' '#1075#1088#1072#1085#1080#1094#1072']'
    Parameters = <>
    Left = 48
    Top = 327
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 96
    Top = 327
    object N1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N2Click
    end
  end
  object ADODataSet2: TADODataSet
    Parameters = <>
    Left = 168
    Top = 327
  end
end
