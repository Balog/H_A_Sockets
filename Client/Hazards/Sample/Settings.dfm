object Setting: TSetting
  Left = 311
  Top = 149
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
  ClientHeight = 694
  ClientWidth = 1016
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Icon.Data = {
    0000010001002020100000000000E80200001600000028000000200000004000
    0000010004000000000080020000000000000000000000000000000000000000
    000000008000008000000080800080000000800080008080000080808000C0C0
    C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF000000
    0000000000000000000000000000FFFFFFFFFFFF88888888FF88BBBB8FFFFFFF
    FFFFFF88888888887BBBBBBBBBB8FFFFFFFFFF8877544577BBBBBBBB8888FFFF
    FFFFF7544444443BBBBBBB8FFFFFFFFFFFF8544444447BBBBBBBB78FFFFFFFFF
    FF8444478883BBBBBBB75778FFFFFFFFF8444788BBBBBBBBB34447878FFFFFFF
    84447FFF8BBBBBB3787444787FFFFFFF5447FFFFFF8778888887444778FFFFF7
    447FFFFFFF877FFF8888744588FFF885448FFF78F87758F878888444788FF884
    47FFFF85788888777F888744788F887747FFFFF78F788F858F888877588F8877
    48FFFFF8FF7FFFF88FF88877488F885548FF8878FF88FFF878888875478F8877
    7775555FF7F7888F77775777578F885448888878F787FFFF45778874488F8877
    48FFFF88F8787FF88FFFF877588F887447FFFFF78F888F87FFFFF844588F8877
    75FFFF84788F88758FFFF775788FF884458FFF778F7788F87FFFF47588FFF887
    7748FFFFFF758FFFFFFF7747F8FFF88844578FFFFF858FFFFFF84758FFFFFF88
    747478FFFF858FFFFF77847FFFFFFFF88744448FFFF58FFF874477FFFFFFFFF8
    887444457887888744447FFFFFFFFFFF88874444444444444448FFFFFFFFFFFF
    F888874444444444478FFFFFFFFFFFFFFFF88887754445778FFFFFFFFFFFFFFF
    FFFF88888888888F8FFFFFFFFFFFFFFFFFFFFFF88888888FFFFFFFFFFFFFFFFF
    FFFF000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poDefault
  OnActivate = FormActivate
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 1016
    Height = 696
    ActivePage = TabSheet2
    TabIndex = 1
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
      object DBGrid1: TDBGrid
        Left = -1
        Top = 0
        Width = 1009
        Height = 668
        DataSource = DataSource1
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
            FieldName = #1053#1072#1079#1074#1072#1085#1080#1077' '#1087#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
            Width = 975
            Visible = True
          end>
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1057#1080#1090#1091#1072#1094#1080#1080
      ImageIndex = 1
      object DBGrid2: TDBGrid
        Left = -1
        Top = 0
        Width = 1009
        Height = 667
        DataSource = DataSource2
        PopupMenu = PopupMenu2
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
    end
  end
  object MainMenu1: TMainMenu
    Left = 176
    Top = 65532
    object N18: TMenuItem
      Caption = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093
      object N19: TMenuItem
        Caption = #1057#1086#1079#1076#1072#1090#1100' '#1085#1086#1074#1091#1102
        OnClick = N19Click
      end
      object N20: TMenuItem
        Caption = #1055#1086#1076#1082#1083#1102#1095#1080#1090#1100#1089#1103
        OnClick = N20Click
      end
      object N21: TMenuItem
        Caption = #1055#1077#1088#1077#1080#1084#1077#1085#1086#1074#1072#1090#1100
        OnClick = N21Click
      end
      object N22: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1082#1072#1082
        OnClick = N22Click
      end
      object N23: TMenuItem
        Caption = #1055#1091#1090#1100' '#1082' '#1073#1072#1079#1077' '#1076#1072#1085#1085#1099#1093
        OnClick = N23Click
      end
    end
    object N4: TMenuItem
      Caption = #1042#1074#1077#1076#1077#1085#1080#1077
      OnClick = N4Click
    end
    object N5: TMenuItem
      Caption = #1040#1083#1075#1086#1088#1080#1090#1084' '#1086#1094#1077#1085#1082#1080
      OnClick = N5Click
    end
    object N6: TMenuItem
      Caption = #1054#1094#1077#1085#1082#1072' '#1079#1085#1072#1095#1080#1084#1086#1089#1090#1080' '#1072#1089#1087#1077#1082#1090#1086#1074
      OnClick = N6Click
    end
    object N17: TMenuItem
      Caption = #1057#1074#1086#1076#1085#1099#1081' '#1088#1077#1077#1089#1090#1088' '#1072#1089#1087#1077#1082#1090#1086#1074
      OnClick = N17Click
    end
    object N3: TMenuItem
      Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
      object N2: TMenuItem
        Caption = #1061#1088#1072#1085#1080#1083#1080#1097#1077' '#1076#1086#1082#1091#1084#1077#1085#1090#1086#1074
        OnClick = N2Click
      end
      object N24: TMenuItem
        Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1082#1088#1080#1090#1077#1088#1080#1077#1074
        OnClick = N24Click
      end
      object N13: TMenuItem
        Caption = #1057#1087#1080#1089#1082#1080
        Enabled = False
      end
      object N14: TMenuItem
        Caption = #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103
        object N15: TMenuItem
          Caption = #1057#1086#1079#1076#1072#1085#1080#1077' '#1082#1083#1102#1095#1077#1074#1086#1075#1086' '#1092#1072#1081#1083#1072
          OnClick = N15Click
        end
        object N16: TMenuItem
          Caption = #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
          OnClick = N16Click
        end
      end
    end
    object N7: TMenuItem
      Caption = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
      OnClick = N7Click
    end
    object N1: TMenuItem
      Caption = #1055#1086#1084#1086#1097#1100
      OnClick = N1Click
    end
    object N8: TMenuItem
      Caption = #1042#1099#1093#1086#1076
      OnClick = N8Click
    end
  end
  object DataSource1: TDataSource
    DataSet = Podr
    Left = 256
  end
  object Podr: TADODataSet
    CursorType = ctStatic
    CommandText = 'select * from '#1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
    Parameters = <>
    Left = 224
  end
  object PopupMenu1: TPopupMenu
    Left = 288
    object N9: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N9Click
    end
    object N10: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N10Click
    end
  end
  object Temp: TADODataSet
    Parameters = <>
    Left = 520
  end
  object Sit: TADODataSet
    CursorType = ctStatic
    CommandText = 'select * from '#1057#1080#1090#1091#1072#1094#1080#1080
    Parameters = <>
    Left = 336
  end
  object DataSource2: TDataSource
    DataSet = Sit
    Left = 368
  end
  object PopupMenu2: TPopupMenu
    Left = 408
    object N11: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N11Click
    end
    object N12: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N12Click
    end
  end
end
