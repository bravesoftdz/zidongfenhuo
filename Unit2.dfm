object Form2: TForm2
  Left = 505
  Top = 267
  BorderStyle = bsSingle
  Caption = #33406#33713#20381#33258#21160#20998#36135#31995#32479
  ClientHeight = 127
  ClientWidth = 337
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 12
  object Label1: TLabel
    Left = 43
    Top = 64
    Width = 246
    Height = 12
    Alignment = taCenter
    Caption = #27491#22312#22788#29702#31532#19968#20010#35746#36135#28165#21333#65292#20849#20116#20010#65292#35831#31245#20505'...'
  end
  object RzProgressBar1: TRzProgressBar
    Left = 16
    Top = 24
    Width = 305
    BorderWidth = 0
    FlatColor = clAqua
    InteriorOffset = 0
    PartsComplete = 0
    Percent = 0
    TotalParts = 0
  end
  object BitBtn1: TBitBtn
    Left = 66
    Top = 88
    Width = 75
    Height = 25
    Caption = #20445#23384
    Enabled = False
    TabOrder = 0
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 200
    Top = 88
    Width = 75
    Height = 25
    Caption = #19979#19968#27493
    Enabled = False
    TabOrder = 1
    OnClick = BitBtn2Click
  end
  object XPManifest1: TXPManifest
    Left = 296
    Top = 88
  end
  object SaveDialog1: TSaveDialog
    Left = 304
    Top = 56
  end
end
