object Report1: TReport1
  Left = 200
  Top = 125
  BorderStyle = bsDialog
  Caption = #1054#1090#1095#1077#1090
  ClientHeight = 143
  ClientWidth = 625
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
  OnKeyUp = FormKeyUp
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 102
    Top = 8
    Width = 124
    Height = 20
    Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1077
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label1: TLabel
    Left = 10
    Top = 32
    Width = 216
    Height = 20
    Caption = #1054#1090#1074#1077#1090#1089#1074#1077#1085#1085#1099#1081' '#1080#1089#1087#1086#1083#1085#1080#1090#1077#1083#1100
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 110
    Top = 56
    Width = 115
    Height = 20
    Caption = #1044#1072#1090#1072' '#1089#1086#1079#1076#1072#1085#1080#1103
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 110
    Top = 80
    Width = 114
    Height = 20
    Caption = #1057#1088#1086#1082' '#1076#1077#1081#1089#1090#1074#1080#1103
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object CPodrazdel: TComboBox
    Left = 232
    Top = 8
    Width = 390
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 0
    OnKeyPress = CPodrazdelKeyPress
  end
  object Edit1: TEdit
    Left = 232
    Top = 32
    Width = 390
    Height = 21
    TabOrder = 1
  end
  object Date1: TDateTimePicker
    Left = 232
    Top = 56
    Width = 186
    Height = 21
    CalAlignment = dtaLeft
    Date = 39917.9508810764
    Time = 39917.9508810764
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 2
    OnKeyPress = Date1KeyPress
  end
  object Date2: TDateTimePicker
    Left = 232
    Top = 80
    Width = 186
    Height = 21
    CalAlignment = dtaLeft
    Date = 39917.9508810764
    Time = 39917.9508810764
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 3
    OnKeyPress = Date2KeyPress
  end
  object Button1: TButton
    Left = 268
    Top = 112
    Width = 89
    Height = 25
    Caption = #1057#1086#1079#1076#1072#1090#1100' '#1086#1090#1095#1077#1090
    TabOrder = 4
    OnClick = Button1Click
  end
  object CheckBox1: TCheckBox
    Left = 423
    Top = 57
    Width = 112
    Height = 17
    Caption = #1053#1072' '#1086#1076#1085#1086#1084' '#1083#1080#1089#1090#1077
    Checked = True
    State = cbChecked
    TabOrder = 5
  end
  object Podr: TADODataSet
    Parameters = <>
    Left = 8
    Top = 8
  end
end
