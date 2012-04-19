object Form1: TForm1
  Left = 192
  Top = 107
  Width = 228
  Height = 207
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
    Width = 201
    Height = 113
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object Edit1: TEdit
    Left = 8
    Top = 128
    Width = 121
    Height = 21
    TabOrder = 1
    Text = 'Edit1'
  end
  object Button1: TButton
    Left = 136
    Top = 128
    Width = 75
    Height = 21
    Caption = 'Button1'
    Enabled = False
    TabOrder = 2
    OnClick = Button1Click
  end
  object Edit2: TEdit
    Left = 8
    Top = 152
    Width = 121
    Height = 21
    TabOrder = 3
    Text = 'Edit2'
  end
  object Button2: TButton
    Left = 136
    Top = 152
    Width = 75
    Height = 21
    Caption = 'Button2'
    Enabled = False
    TabOrder = 4
    OnClick = Button2Click
  end
  object ServerSocket1: TServerSocket
    Active = False
    Port = 1001
    ServerType = stNonBlocking
    OnClientConnect = ServerSocket1ClientConnect
    OnClientDisconnect = ServerSocket1ClientDisconnect
    OnClientRead = ServerSocket1ClientRead
    Left = 16
    Top = 24
  end
  object ActionManager1: TActionManager
    Left = 56
    Top = 24
    object Comm1: TAction
      Caption = 'Comm1'
      OnExecute = Comm1Execute
    end
    object Comm2: TAction
      Caption = 'Comm2'
      OnExecute = Comm2Execute
    end
    object Comm3: TAction
      Caption = 'Comm3'
      OnExecute = Comm3Execute
    end
  end
end
