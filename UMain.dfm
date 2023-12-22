object Form1: TForm1
  Left = 206
  Top = 694
  Caption = 'Importfis'
  ClientHeight = 536
  ClientWidth = 1000
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  WindowState = wsMaximized
  OnCreate = FormCreate
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label4: TLabel
    Left = 32
    Top = 8
    Width = 31
    Height = 13
    Caption = 'Label4'
  end
  object ProgressBar1: TProgressBar
    Left = 0
    Top = 519
    Width = 1000
    Height = 17
    Align = alBottom
    TabOrder = 0
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 1000
    Height = 519
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = 'Home'
      object PanelUnder: TPanel
        Left = 0
        Top = 108
        Width = 992
        Height = 383
        Align = alClient
        TabOrder = 0
        object Panel1: TPanel
          Left = 1
          Top = 1
          Width = 990
          Height = 56
          Align = alTop
          TabOrder = 0
          object Label5: TLabel
            Left = 19
            Top = 5
            Width = 5
            Height = 19
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clSilver
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Panel4: TPanel
            Left = 720
            Top = 1
            Width = 269
            Height = 54
            Align = alRight
            TabOrder = 0
            object Label6: TLabel
              Left = 61
              Top = 0
              Width = 192
              Height = 19
              Caption = 'Selecione coluna where'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -16
              Font.Name = 'Tahoma'
              Font.Style = [fsBold]
              ParentFont = False
            end
            object ComboBox1: TComboBox
              Left = 61
              Top = 25
              Width = 193
              Height = 22
              Style = csOwnerDrawVariable
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clRed
              Font.Height = -12
              Font.Name = 'Tahoma'
              Font.Pitch = fpFixed
              Font.Style = [fsBold]
              ParentFont = False
              TabOrder = 0
              Items.Strings = (
                'Selecione...'
                'COD_ITEM'
                'COD_BARRA'
                'COD'
                'EAN')
            end
          end
          object CheckBox1: TCheckBox
            Left = 2
            Top = 30
            Width = 174
            Height = 17
            Caption = 'Remover pontua'#231#227'o?'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
            TabOrder = 1
          end
        end
        object ScrollBox2: TScrollBox
          Left = 1
          Top = 57
          Width = 990
          Height = 325
          Align = alClient
          TabOrder = 1
          object StringGrid1: TStringGrid
            Left = 0
            Top = 0
            Width = 632
            Height = 321
            Align = alClient
            DefaultColWidth = 48
            DrawingStyle = gdsClassic
            FixedCols = 0
            FixedRows = 2
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 0
            OnMouseDown = StringGrid1MouseDown
            OnMouseMove = StringGrid1MouseMove
            ColWidths = (
              48
              48
              48
              48
              48)
          end
          object Panel2: TPanel
            Left = 632
            Top = 0
            Width = 354
            Height = 321
            Align = alRight
            TabOrder = 1
            object ValueListEditor1: TValueListEditor
              Left = 1
              Top = 36
              Width = 352
              Height = 284
              Align = alClient
              DrawingStyle = gdsClassic
              FixedColor = clMenu
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clRed
              Font.Height = -11
              Font.Name = 'Tahoma'
              Font.Style = [fsBold]
              ParentFont = False
              ScrollBars = ssHorizontal
              TabOrder = 0
              ColWidths = (
                150
                196)
            end
            object Panel5: TPanel
              Left = 1
              Top = 1
              Width = 352
              Height = 35
              Align = alTop
              TabOrder = 1
              object BitBtn2: TBitBtn
                Left = 13
                Top = 5
                Width = 124
                Height = 25
                Align = alCustom
                Caption = 'Limpar tudo'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clRed
                Font.Height = -13
                Font.Name = 'Tahoma'
                Font.Style = []
                ParentFont = False
                Style = bsNew
                TabOrder = 0
                OnClick = BitBtn2Click
              end
              object BitBtn3: TBitBtn
                Left = 157
                Top = 5
                Width = 124
                Height = 25
                Align = alCustom
                Caption = 'apagar registro'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clRed
                Font.Height = -13
                Font.Name = 'Tahoma'
                Font.Style = []
                ParentFont = False
                Style = bsNew
                TabOrder = 1
                OnClick = BitBtn3Click
              end
            end
          end
        end
      end
      object pnlTop: TPanel
        Left = 0
        Top = 0
        Width = 992
        Height = 108
        Align = alTop
        TabOrder = 1
        DesignSize = (
          992
          108)
        object Label2: TLabel
          Left = 40
          Top = 49
          Width = 66
          Height = 19
          Caption = 'Planilha'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Font.Quality = fqDraft
          ParentFont = False
        end
        object Label1: TLabel
          Left = 20
          Top = 8
          Width = 86
          Height = 19
          Caption = 'Banco FDB'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          Font.Quality = fqDraft
          ParentFont = False
        end
        object btnImport: TSpeedButton
          Left = 682
          Top = 75
          Width = 100
          Height = 28
          Anchors = [akTop, akRight]
          Caption = 'Importar'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          OnClick = btnImportClick
        end
        object Label3: TLabel
          Left = 20
          Top = 77
          Width = 510
          Height = 19
          Caption = 'Clique no nome da coluna e selecione o campo correspondente'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object btnLoadExcel: TSpeedButton
          Left = 576
          Top = 75
          Width = 100
          Height = 28
          Anchors = [akTop, akRight]
          Caption = 'Carregar planilha'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          OnClick = btnLoadExcelClick
        end
        object SpeedButton2: TSpeedButton
          Left = 890
          Top = 74
          Width = 100
          Height = 28
          Anchors = [akTop, akRight]
          Caption = 'Fechar'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          OnClick = SpeedButton2Click
        end
        object SpeedButton1: TSpeedButton
          Left = 784
          Top = 75
          Width = 100
          Height = 28
          Anchors = [akTop, akRight]
          Caption = 'Importar'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          OnClick = SpeedButton1Click
        end
        object edtfdb: TEdit
          Left = 128
          Top = 5
          Width = 807
          Height = 27
          Anchors = [akLeft, akTop, akRight]
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          Text = 'C:\DB_SAT_FACIL.FDB'
        end
        object edtExcel: TEdit
          Left = 128
          Top = 38
          Width = 807
          Height = 27
          Anchors = [akLeft, akTop, akRight]
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          Text = 'C:\CadastroItens (1).xls'
        end
        object BitBtn1: TBitBtn
          Left = 941
          Top = 38
          Width = 49
          Height = 27
          Anchors = [akTop, akRight]
          Caption = '...'
          TabOrder = 2
          OnClick = BitBtn1Click
        end
        object btnConectar: TBitBtn
          Left = 941
          Top = 5
          Width = 49
          Height = 27
          Anchors = [akTop, akRight]
          Caption = '...'
          ModalResult = 5
          NumGlyphs = 2
          TabOrder = 3
          OnClick = btnConectarClick
        end
      end
    end
    object ts_log: TTabSheet
      Caption = 'Log'
      ImageIndex = 1
      object Panel3: TPanel
        Left = 0
        Top = 450
        Width = 992
        Height = 41
        Align = alBottom
        Caption = 'Log'
        TabOrder = 0
      end
      object Memo1: TMemo
        Left = 0
        Top = 0
        Width = 992
        Height = 450
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 1
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Update List'
      ImageIndex = 2
      object memo2: TMemo
        Left = 0
        Top = 0
        Width = 992
        Height = 491
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGreen
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 0
      end
    end
  end
  object FDBcoERP: TFDConnection
    Params.Strings = (
      'DriverID=FB'
      'Database=C:\desafio\DESAFIO.FDB'
      'Password=masterkey'
      'User_Name=SYSDBA'
      'Port=3050')
    LoginPrompt = False
    Left = 16
    Top = 48
  end
  object qryFieldProduto: TFDQuery
    Connection = FDBcoERP
    SQL.Strings = (
      'SELECT RDB$FIELD_NAME AS FIELD'
      'FROM RDB$RELATION_FIELDS'
      'WHERE RDB$RELATION_NAME = '#39'PRODUTOS'#39
      'ORDER BY RDB$FIELD_NAME;')
    Left = 24
    Top = 264
  end
  object FDGUIxWaitCursor1: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 32
    Top = 216
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    Left = 264
    Top = 80
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = 'csv'
    Filter = 'csv'
    Left = 544
    Top = 128
  end
  object FDQueryProduto: TFDQuery
    Connection = FDBcoERP
    SQL.Strings = (
      'select * from produtos p'
      'where'
      'p.codigo_barras=:produto')
    Left = 184
    Top = 264
    ParamData = <
      item
        Name = 'PRODUTO'
        ParamType = ptInput
      end>
  end
end
