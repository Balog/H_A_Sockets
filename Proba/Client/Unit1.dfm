object Form1: TForm1
  Left = 422
  Top = 103
  Width = 321
  Height = 160
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Memo1: TMemo
    Left = 8
    Top = 8
    Width = 297
    Height = 89
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object Button1: TButton
    Left = 8
    Top = 104
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 1
    OnClick = Button1Click
  end
  object ClientSocket1: TClientSocket
    Active = False
    Address = 'localhost'
    ClientType = ctNonBlocking
    Host = 'localhost'
    Port = 1001
    OnRead = ClientSocket1Read
    Left = 264
    Top = 16
  end
  object ActionManager1: TActionManager
    Left = 96
    Top = 104
    object Command1: TAction
      Caption = 'Command1'
      OnExecute = Command1Execute
    end
    object Command2: TAction
      Caption = 'Command2'
      OnExecute = Command2Execute
    end
    object Command3: TAction
      Caption = 'Command3'
      OnExecute = Command3Execute
    end
  end
end
