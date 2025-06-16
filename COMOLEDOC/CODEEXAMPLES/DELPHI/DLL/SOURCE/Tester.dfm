object FormTester: TFormTester
  Left = 441
  Top = 0
  BorderStyle = bsSingle
  Caption = 'FormCaption'
  ClientHeight = 522
  ClientWidth = 671
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  DesignSize = (
    671
    522)
  PixelsPerInch = 96
  TextHeight = 13
  object LabelSymbols: TLabel
    Left = 8
    Top = 8
    Width = 39
    Height = 13
    Caption = 'Symbols'
  end
  object LabelLines: TLabel
    Left = 8
    Top = 253
    Width = 24
    Height = 13
    Caption = 'Lines'
  end
  object ListBoxAllSymbols: TListBox
    Left = 8
    Top = 22
    Width = 655
    Height = 217
    Anchors = [akLeft, akTop, akRight]
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Lucida Console'
    Font.Style = []
    ItemHeight = 11
    ParentFont = False
    TabOrder = 0
  end
  object ListBoxAllLines: TListBox
    Left = 8
    Top = 267
    Width = 655
    Height = 217
    Anchors = [akLeft, akTop, akRight]
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Lucida Console'
    Font.Style = []
    ItemHeight = 11
    ParentFont = False
    ScrollWidth = 1000
    TabOrder = 1
  end
  object ButtonGetSymbolsAndLines: TButton
    Left = 8
    Top = 490
    Width = 121
    Height = 25
    Caption = 'Get symbols and lines'
    TabOrder = 2
    OnClick = ButtonGetSymbolsAndLinesClick
  end
end
