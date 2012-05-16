object FDiary: TFDiary
  Left = 192
  Top = 107
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1046#1091#1088#1085#1072#1083' '#1089#1086#1073#1099#1090#1080#1081
  ClientHeight = 713
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
  OnCreate = FormCreate
  OnKeyUp = FormKeyUp
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 35
    Top = 3
    Width = 81
    Height = 13
    Caption = #1053#1072#1095#1072#1083#1100#1085#1072#1103' '#1076#1072#1090#1072
  end
  object Label2: TLabel
    Left = 205
    Top = 3
    Width = 74
    Height = 13
    Caption = #1050#1086#1085#1077#1095#1085#1072#1103' '#1076#1072#1090#1072
  end
  object Label3: TLabel
    Left = 365
    Top = 3
    Width = 66
    Height = 13
    Caption = #1050#1086#1084#1087#1100#1102#1090#1077#1088#1099
  end
  object Label4: TLabel
    Left = 537
    Top = 3
    Width = 39
    Height = 13
    Caption = #1051#1086#1075#1080#1085#1099
  end
  object Label5: TLabel
    Left = 769
    Top = 3
    Width = 87
    Height = 13
    Caption = #1058#1080#1087#1099' '#1089#1086#1086#1073#1097#1077#1085#1080#1081
  end
  object NDate: TMonthCalendar
    Left = 3
    Top = 21
    Width = 155
    Height = 154
    Date = 40897.3916392361
    TabOrder = 0
    OnClick = NDateClick
  end
  object KDate: TMonthCalendar
    Left = 165
    Top = 21
    Width = 155
    Height = 154
    Date = 40897.3916392361
    TabOrder = 1
    OnClick = NDateClick
  end
  object Comps: TCheckListBox
    Left = 325
    Top = 21
    Width = 153
    Height = 153
    OnClickCheck = CompsClickCheck
    ItemHeight = 13
    TabOrder = 2
  end
  object Logins: TCheckListBox
    Left = 483
    Top = 21
    Width = 153
    Height = 153
    OnClickCheck = CompsClickCheck
    ItemHeight = 13
    TabOrder = 3
  end
  object Types: TCheckListBox
    Left = 642
    Top = 21
    Width = 371
    Height = 153
    OnClickCheck = CompsClickCheck
    ItemHeight = 13
    TabOrder = 4
  end
  object Diary: TDBGrid
    Left = 0
    Top = 183
    Width = 1014
    Height = 530
    DataSource = DataSource1
    TabOrder = 5
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'Date_Time'
        Title.Caption = #1044#1072#1090#1072'/'#1042#1088#1077#1084#1103
        Width = 107
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Comp'
        Title.Caption = #1050#1086#1084#1087#1100#1102#1090#1077#1088
        Width = 91
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Login'
        Title.Caption = #1051#1086#1075#1080#1085
        Width = 83
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NameType'
        Title.Caption = #1058#1080#1087' '#1089#1086#1086#1073#1097#1077#1085#1080#1103
        Width = 141
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NameOperation'
        Title.Caption = #1057#1086#1086#1073#1097#1077#1085#1080#1077
        Width = 310
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Prim'
        Title.Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
        Width = 243
        Visible = True
      end>
  end
  object EnNDate: TCheckBox
    Left = 4
    Top = -1
    Width = 17
    Height = 17
    Checked = True
    State = cbChecked
    TabOrder = 6
    OnClick = EnNDateClick
  end
  object EnKDate: TCheckBox
    Left = 166
    Top = -1
    Width = 17
    Height = 17
    Checked = True
    State = cbChecked
    TabOrder = 7
    OnClick = EnKDateClick
  end
  object Button1: TButton
    Left = 950
    Top = 0
    Width = 62
    Height = 19
    Caption = #1054#1073#1085#1086#1074#1080#1090#1100
    TabOrder = 8
    OnClick = Button1Click
  end
  object PB: TProgressBar
    Left = 2
    Top = 175
    Width = 1011
    Height = 6
    Min = 0
    Max = 6
    Smooth = True
    TabOrder = 9
    Visible = False
  end
  object Button2: TButton
    Left = 885
    Top = 1
    Width = 62
    Height = 19
    Caption = #1042' Excel'
    TabOrder = 10
    OnClick = Button2Click
  end
  object DataSource1: TDataSource
    DataSet = Events
    Left = 336
    Top = 240
  end
  object Events: TADODataSet
    CursorType = ctStatic
    Parameters = <>
    Left = 376
    Top = 240
  end
end
