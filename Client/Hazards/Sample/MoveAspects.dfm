object MoveAsp: TMoveAsp
  Left = 192
  Top = 107
  Width = 1024
  Height = 740
  Caption = 'MoveAsp'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object TabAsp: TStringGrid
    Left = 0
    Top = 0
    Width = 1016
    Height = 713
    Align = alClient
    ColCount = 6
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    TabOrder = 0
    OnSelectCell = TabAspSelectCell
    ColWidths = (
      153
      162
      158
      146
      152
      59)
  end
end
