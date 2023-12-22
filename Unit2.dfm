object Form2: TForm2
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  Caption = 'Selecione o Campo da tabela '
  ClientHeight = 107
  ClientWidth = 631
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIForm
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 631
    Height = 41
    Align = alTop
    Caption = 'Selecione...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
  end
  object Panel2: TPanel
    Left = 0
    Top = 66
    Width = 631
    Height = 41
    Align = alBottom
    TabOrder = 1
    object btnOK: TBitBtn
      Left = 1
      Top = 1
      Width = 629
      Height = 39
      Align = alClient
      Caption = 'Confirma'
      ModalResult = 1
      TabOrder = 0
      OnClick = btnOKClick
    end
  end
  object ComboBox1: TComboBox
    Left = 0
    Top = 41
    Width = 631
    Height = 36
    Align = alClient
    AutoDropDown = True
    Style = csOwnerDrawVariable
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -20
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ItemHeight = 30
    ItemIndex = 0
    ParentFont = False
    TabOrder = 2
    Text = 'Selecione...'
    Items.Strings = (
      'Selecione...')
  end
end
