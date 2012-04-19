object Filter: TFilter
  Left = 292
  Top = 197
  BorderStyle = bsDialog
  Caption = #1060#1080#1083#1100#1090#1088
  ClientHeight = 296
  ClientWidth = 828
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
    C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FFFF
    FFFFFFFFFF8888FFFFFFFFFFFFFFFFFFFFFFFFF8888777888FFFFFFFFFFFFFFF
    FFFF8887444477477788FFFFFFFFFFFFFFF887777844878777778FFFFFFFFFFF
    F887748777447474777778FFFFFFFFFFF877774478F88F877477478FFFFFFFFF
    8748747FFF8748FFF8448748FFFFFFF8777778FF85444447FF8774778FFFFFF8
    478447747FF448F74777487478FFFF877544787888F87F77F787447878FFF887
    7776F77FF857777FF84F747747FFF8747784F8878F7444F8788F7774778FF878
    77F4FFF77884448778FF7887878FF87878F87FFF88F448F7FFF87F87777F8844
    48FF77777FF448F77777FFF4447887444FFFFF88FFF448FFF88FFFF444788747
    4FFFF7778F7444887778FFF47478884448FF7887787877777784FFF444788847
    78FF4F7774887F7777F78FF4747FF84777FF7FF778F448F778F78F87748FF877
    74FF77FFF874447FFF84FF77748FFF74787FF48FFFF748FFF848F87747FFFF84
    7748FF7778874777778FF74848FFFFF74775FFFF788748888FFF774778FFFFF8
    774778FFFF8577FFFFF777748FFFFFFF8478748FFF8777FFF8778748FFFFFFFF
    F847487788F8788877477488FFFFFFFFFF877877777747477774788FFFFFFFFF
    FFF8747787884777844788FFFFFFFFFFFFFF8877477747745788FFFFFFFFFFFF
    FFFFFF888787777888FFFFFFFFFFFFFFFFFFFFFFFF8888FFFFFFFFFFFFFF0000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  KeyPreview = True
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnKeyUp = FormKeyUp
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object RadioGroup1: TRadioGroup
    Left = 8
    Top = 5
    Width = 817
    Height = 249
    Caption = #1042#1099#1073#1086#1088' '#1092#1080#1083#1100#1090#1088#1072
    ItemIndex = 0
    Items.Strings = (
      #1054#1090#1082#1083#1102#1095#1080#1090#1100' '#1092#1080#1083#1100#1090#1088
      #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1077
      #1042#1080#1076' '#1090#1077#1088#1088#1080#1090#1086#1088#1080#1080
      #1042#1080#1076' '#1076#1077#1103#1090#1077#1083#1100#1085#1086#1089#1090#1080', '#1086#1087#1077#1088#1072#1094#1080#1103
      #1040#1089#1087#1077#1082#1090
      #1042#1086#1079#1076#1077#1081#1089#1090#1074#1080#1077
      #1042#1072#1083#1080#1076#1085#1086#1089#1090#1100
      #1047#1085#1072#1095#1080#1084#1086#1089#1090#1100)
    TabOrder = 0
    OnClick = RadioGroup1Click
  end
  object Edit1: TEdit
    Left = 184
    Top = 81
    Width = 594
    Height = 21
    ReadOnly = True
    TabOrder = 1
    Visible = False
    OnDblClick = Edit1DblClick
  end
  object Edit2: TEdit
    Left = 184
    Top = 108
    Width = 594
    Height = 21
    ReadOnly = True
    TabOrder = 2
    Visible = False
    OnDblClick = Edit2DblClick
  end
  object Edit3: TEdit
    Left = 184
    Top = 135
    Width = 594
    Height = 21
    ReadOnly = True
    TabOrder = 3
    Visible = False
    OnDblClick = Edit3DblClick
  end
  object Edit4: TEdit
    Left = 184
    Top = 162
    Width = 594
    Height = 21
    ReadOnly = True
    TabOrder = 4
    Visible = False
    OnDblClick = Edit4DblClick
  end
  object Button1: TButton
    Left = 784
    Top = 81
    Width = 33
    Height = 21
    Caption = '...'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 5
    Visible = False
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 784
    Top = 108
    Width = 33
    Height = 21
    Caption = '...'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 6
    Visible = False
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 784
    Top = 135
    Width = 33
    Height = 21
    Caption = '...'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 7
    Visible = False
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 784
    Top = 162
    Width = 33
    Height = 21
    Caption = '...'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 8
    Visible = False
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 376
    Top = 264
    Width = 75
    Height = 25
    Caption = #1042#1074#1086#1076
    Enabled = False
    TabOrder = 9
    OnClick = Button5Click
  end
  object ComboBox1: TComboBox
    Left = 184
    Top = 54
    Width = 290
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 10
    Visible = False
    OnClick = ComboBox1Click
    OnKeyPress = ComboBox1KeyPress
  end
  object ComboBox2: TComboBox
    Left = 184
    Top = 190
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 11
    Visible = False
    OnClick = ComboBox2Click
    OnKeyPress = ComboBox2KeyPress
    Items.Strings = (
      #1042#1072#1083#1080#1076#1085#1099#1077
      #1053#1077#1074#1072#1083#1080#1076#1085#1099#1077)
  end
  object ComboBox3: TComboBox
    Left = 184
    Top = 218
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 12
    Visible = False
    OnClick = ComboBox3Click
    OnKeyPress = ComboBox3KeyPress
    Items.Strings = (
      #1047#1085#1072#1095#1080#1084#1099#1077
      #1053#1077#1079#1085#1072#1095#1080#1084#1099#1077)
  end
  object Podr: TADODataSet
    CursorType = ctStatic
    CommandText = 'select * from '#1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103
    Parameters = <>
    Left = 480
    Top = 46
  end
  object DataSource1: TDataSource
    DataSet = Podr
    Left = 520
    Top = 46
  end
end
