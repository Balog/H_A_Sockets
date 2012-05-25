object Documents: TDocuments
  Left = 209
  Top = 136
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #1057#1087#1088#1072#1074#1086#1095#1085#1080#1082#1080
  ClientHeight = 736
  ClientWidth = 1031
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 1031
    Height = 736
    ActivePage = TabPodr
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MultiLine = True
    ParentFont = False
    TabIndex = 1
    TabOrder = 0
    object TabMetod: TTabSheet
      Caption = #1052#1077#1090#1086#1076#1080#1082#1072
      ImageIndex = 8
      object DBMemo1: TDBMemo
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        DataField = #1052#1077#1090#1086#1076#1080#1082#1072
        DataSource = DataSource5
        Font.Charset = RUSSIAN_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Courier New'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
      end
      object Panel1: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object ReadMetod: TCheckBox
          Left = 72
          Top = 4
          Width = 97
          Height = 17
          Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1090#1100
          TabOrder = 0
        end
        object Button1: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 1
        end
      end
    end
    object TabPodr: TTabSheet
      Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
      ImageIndex = 6
      object DBGrid2: TDBGrid
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        DataSource = DataSource2
        PopupMenu = PopupMenu7
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = #1053#1072#1079#1074#1072#1085#1080#1077' '#1087#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
            Width = 975
            Visible = True
          end>
      end
      object Panel2: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button2: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabCryt: TTabSheet
      Caption = #1050#1088#1080#1090#1077#1088#1080#1080
      ImageIndex = 5
      object DBGrid1: TDBGrid
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        DataSource = DataSource1
        Options = [dgEditing, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs]
        PopupMenu = PopupMenu6
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
      object Panel3: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button3: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabSit: TTabSheet
      Caption = #1057#1080#1090#1091#1072#1094#1080#1080
      ImageIndex = 7
      object DBGrid3: TDBGrid
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        DataSource = DataSource3
        PopupMenu = PopupMenu8
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = #1053#1072#1079#1074#1072#1085#1080#1077' '#1089#1080#1090#1091#1072#1094#1080#1080
            Width = 974
            Visible = True
          end>
      end
      object Panel9: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button9: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabVozd: TTabSheet
      Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1074#1086#1079#1076#1077#1081#1089#1090#1074#1080#1103
      object TreeView1: TTreeView
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Indent = 19
        ParentFont = False
        TabOrder = 0
        ToolTips = False
      end
      object Panel4: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button4: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabMeropr: TTabSheet
      Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1087#1088#1080#1088#1086#1076#1086#1086#1093#1088#1072#1085#1085#1099#1093' '#1084#1077#1088#1086#1087#1088#1080#1103#1090#1080#1081
      ImageIndex = 1
      object TreeView2: TTreeView
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Indent = 19
        ParentFont = False
        TabOrder = 0
      end
      object Panel5: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button5: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabTerr: TTabSheet
      Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1090#1077#1088#1088#1080#1090#1086#1088#1080#1081
      ImageIndex = 2
      object TreeView3: TTreeView
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Indent = 19
        ParentFont = False
        TabOrder = 0
      end
      object Panel6: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button6: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabDeyat: TTabSheet
      Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080
      ImageIndex = 3
      object TreeView4: TTreeView
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Indent = 19
        ParentFont = False
        TabOrder = 0
      end
      object Panel7: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button7: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
    object TabAspects: TTabSheet
      Caption = #1054#1073#1097#1080#1081' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082' '#1101#1082#1086#1083#1086#1075#1080#1095#1077#1089#1082#1080#1093' '#1072#1089#1087#1077#1082#1090#1086#1074
      ImageIndex = 4
      object TreeView5: TTreeView
        Left = 0
        Top = 0
        Width = 1023
        Height = 666
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Indent = 19
        ParentFont = False
        TabOrder = 0
      end
      object Panel8: TPanel
        Left = 0
        Top = 666
        Width = 1023
        Height = 24
        Align = alBottom
        TabOrder = 1
        object Button8: TButton
          Left = 5
          Top = 4
          Width = 61
          Height = 19
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
          TabOrder = 0
        end
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 312
    Top = 129
    object N4: TMenuItem
      Caption = #1063#1090#1077#1085#1080#1077
      object N2: TMenuItem
        Caption = #1052#1077#1090#1086#1076#1080#1082#1072
      end
      object N27: TMenuItem
        Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
      end
      object N26: TMenuItem
        Caption = #1050#1088#1080#1090#1077#1088#1080#1080
      end
      object N28: TMenuItem
        Caption = #1057#1080#1090#1091#1072#1094#1080#1080
      end
      object N6: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1074#1086#1079#1076#1077#1081#1089#1090#1074#1080#1103
      end
      object N7: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1087#1088#1080#1088#1086#1076#1086#1086#1093#1088#1072#1085#1085#1099#1093' '#1084#1077#1088#1086#1087#1088#1080#1103#1090#1080#1081
      end
      object N20: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1090#1077#1088#1088#1080#1090#1086#1088#1080#1081
      end
      object N24: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080
      end
      object N25: TMenuItem
        Caption = #1054#1073#1097#1080#1081' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082' '#1101#1082#1086#1083#1086#1075#1080#1095#1077#1089#1082#1080#1093' '#1072#1089#1087#1077#1082#1090#1086#1074
      end
      object N29: TMenuItem
        Caption = '-'
      end
      object N31: TMenuItem
        Caption = #1055#1088#1086#1095#1080#1090#1072#1090#1100' '#1074#1089#1077' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1080
      end
    end
    object N5: TMenuItem
      Caption = #1047#1072#1087#1080#1089#1100
      object N18: TMenuItem
        Caption = #1052#1077#1090#1086#1076#1080#1082#1072
      end
      object N38: TMenuItem
        Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
      end
      object N37: TMenuItem
        Caption = #1050#1088#1080#1090#1077#1088#1080#1080
      end
      object N39: TMenuItem
        Caption = #1057#1080#1090#1091#1072#1094#1080#1080
      end
      object N32: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1074#1086#1079#1076#1077#1081#1089#1090#1074#1080#1103
      end
      object N33: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1087#1088#1080#1088#1086#1076#1086#1086#1093#1088#1072#1085#1085#1099#1093' '#1084#1077#1088#1086#1087#1088#1080#1103#1090#1080#1081
      end
      object N34: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1090#1077#1088#1088#1080#1090#1086#1088#1080#1081
      end
      object N35: TMenuItem
        Caption = #1055#1077#1088#1077#1095#1077#1085#1100' '#1074#1080#1076#1086#1074' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080
      end
      object N36: TMenuItem
        Caption = #1054#1073#1097#1080#1081' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082' '#1101#1082#1086#1083#1086#1075#1080#1095#1077#1089#1082#1080#1093' '#1072#1089#1087#1077#1082#1090#1086#1074
      end
      object N40: TMenuItem
        Caption = '-'
      end
      object N41: TMenuItem
        Caption = #1047#1072#1087#1080#1089#1072#1090#1100' '#1074#1089#1077' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1080
      end
    end
    object N19: TMenuItem
      Caption = #1044#1074#1080#1078#1077#1085#1080#1077' '#1072#1089#1087#1077#1082#1090#1086#1074
    end
    object N21: TMenuItem
      Caption = #1054#1090#1095#1077#1090#1099
      object N00111: TMenuItem
        Caption = #1060'001.1'
      end
      object N00121: TMenuItem
        Caption = #1060'001.2'
      end
      object N22: TMenuItem
        Caption = #1057#1074#1086#1076#1085#1099#1081' '#1086#1090#1095#1077#1090
      end
    end
    object N1: TMenuItem
      Caption = #1055#1086#1084#1086#1097#1100
    end
    object N23: TMenuItem
      Caption = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
    end
    object N8: TMenuItem
      Caption = #1042#1099#1093#1086#1076
    end
  end
  object Nodes: TADODataSet
    Parameters = <>
    Left = 496
    Top = 624
  end
  object Branches: TADODataSet
    Parameters = <>
    Left = 536
    Top = 624
  end
  object Comm: TADOCommand
    Parameters = <>
    Left = 576
    Top = 624
  end
  object PopupMenu5: TPopupMenu
    Left = 624
    Top = 624
    object N16: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N16Click
    end
    object N17: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N17Click
    end
  end
  object PopupMenu4: TPopupMenu
    Left = 624
    Top = 592
    object N14: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N14Click
    end
    object N15: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N15Click
    end
  end
  object PopupMenu3: TPopupMenu
    Left = 624
    Top = 560
    object N12: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N12Click
    end
    object N13: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N13Click
    end
  end
  object PopupMenu2: TPopupMenu
    Left = 624
    Top = 528
    object N10: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N10Click
    end
    object N11: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N11Click
    end
  end
  object PopupMenu1: TPopupMenu
    Left = 624
    Top = 496
    object N3: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N3Click
    end
    object N9: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N9Click
    end
  end
  object Podr: TADODataSet
    CursorType = ctStatic
    CommandText = 'select * from '#1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
    Parameters = <>
    Left = 704
    Top = 584
  end
  object DataSource2: TDataSource
    DataSet = Podr
    Left = 736
    Top = 584
  end
  object PopupMenu7: TPopupMenu
    Left = 768
    Top = 584
    object MenuItem3: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = MenuItem3Click
    end
    object MenuItem4: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = MenuItem4Click
    end
  end
  object DataSource1: TDataSource
    DataSet = ADODataSet1
    Left = 816
    Top = 543
  end
  object ADODataSet1: TADODataSet
    CursorType = ctStatic
    CommandText = 'select * from '#1047#1085#1072#1095#1080#1084#1086#1089#1090#1100' Order By ['#1052#1072#1082#1089' '#1075#1088#1072#1085#1080#1094#1072']'
    Parameters = <>
    Left = 856
    Top = 543
  end
  object Sit: TADODataSet
    CursorType = ctStatic
    CommandText = 'SELECT '#1057#1080#1090#1091#1072#1094#1080#1080'.* FROM '#1057#1080#1090#1091#1072#1094#1080#1080#13#10
    Parameters = <>
    Left = 856
    Top = 584
  end
  object DataSource3: TDataSource
    DataSet = Sit
    Left = 888
    Top = 584
  end
  object Metod: TADODataSet
    CursorType = ctStatic
    CommandText = 'select * from '#1052#1077#1090#1086#1076#1080#1082#1072
    Parameters = <>
    Left = 888
    Top = 624
  end
  object PopupMenu6: TPopupMenu
    Left = 904
    Top = 543
    object MenuItem1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = MenuItem1Click
    end
    object MenuItem2: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = MenuItem2Click
    end
  end
  object PopupMenu8: TPopupMenu
    Left = 928
    Top = 584
    object MenuItem5: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = MenuItem5Click
    end
    object MenuItem6: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = MenuItem6Click
    end
  end
  object DataSource5: TDataSource
    DataSet = Metod
    Left = 928
    Top = 624
  end
  object ActionManager1: TActionManager
    Left = 976
    Top = 624
    object DataSetPost1: TDataSetPost
      Category = 'Dataset'
      Caption = 'P&ost'
      Hint = 'Post'
      ImageIndex = 7
      DataSource = DataSource5
    end
    object DataSetPost2: TDataSetPost
      Category = 'Dataset'
      Caption = 'P&ost'
      Hint = 'Post'
      ImageIndex = 7
      DataSource = DataSource1
    end
    object DataSetPost3: TDataSetPost
      Category = 'Dataset'
      Caption = 'P&ost'
      Hint = 'Post'
      ImageIndex = 7
      DataSource = DataSource2
    end
    object DataSetPost4: TDataSetPost
      Category = 'Dataset'
      Caption = 'P&ost'
      Hint = 'Post'
      ImageIndex = 7
      DataSource = DataSource3
    end
    object DataSetRefresh1: TDataSetRefresh
      Category = 'Dataset'
      Caption = '&Refresh'
      Hint = 'Refresh'
      ImageIndex = 9
      DataSource = DataSource3
    end
    object DataSetRefresh2: TDataSetRefresh
      Category = 'Dataset'
      Caption = '&Refresh'
      Hint = 'Refresh'
      ImageIndex = 9
      DataSource = DataSource1
    end
    object DataSetRefresh3: TDataSetRefresh
      Category = 'Dataset'
      Caption = '&Refresh'
      Hint = 'Refresh'
      ImageIndex = 9
      DataSource = DataSource2
    end
    object DataSetRefresh4: TDataSetRefresh
      Category = 'Dataset'
      Caption = '&Refresh'
      Hint = 'Refresh'
      ImageIndex = 9
      DataSource = DataSource5
    end
  end
  object Temp: TADODataSet
    Parameters = <>
    Left = 976
    Top = 584
  end
  object ADODataSet2: TADODataSet
    Parameters = <>
    Left = 976
    Top = 543
  end
end
