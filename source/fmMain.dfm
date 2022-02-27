object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'Directory Builder'
  ClientHeight = 531
  ClientWidth = 988
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  DesignSize = (
    988
    531)
  TextHeight = 15
  object Button1: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Load'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 89
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Generate'
    TabOrder = 1
    OnClick = Button2Click
  end
  object CheckListBox1: TCheckListBox
    Left = 8
    Top = 39
    Width = 457
    Height = 484
    Anchors = [akLeft, akTop, akBottom]
    ItemHeight = 15
    TabOrder = 2
  end
  object Memo1: TMemo
    Left = 471
    Top = 39
    Width = 509
    Height = 484
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 3
  end
end
