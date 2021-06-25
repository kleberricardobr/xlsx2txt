object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 365
  ClientWidth = 447
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object lbStatus: TLabel
    Left = 56
    Top = 304
    Width = 83
    Height = 24
    Caption = 'lbStatus'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -20
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Button2: TButton
    Left = 184
    Top = 246
    Width = 75
    Height = 25
    Caption = 'Load'
    TabOrder = 0
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 8
    Top = 246
    Width = 153
    Height = 25
    Caption = 'LoadXLSX'
    TabOrder = 1
    OnClick = Button3Click
  end
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 447
    Height = 185
    Align = alTop
    FixedCols = 0
    FixedRows = 0
    TabOrder = 2
  end
  object Button1: TButton
    Left = 8
    Top = 191
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 3
    OnClick = Button1Click
  end
  object OpenDialog1: TOpenDialog
    Left = 392
    Top = 240
  end
  object odXLS: TOpenDialog
    Left = 384
    Top = 296
  end
  object qImportaPlan: TJvMemoryData
    Active = True
    FieldDefs = <
      item
        Name = 'dt_pedido'
        DataType = ftDateTime
      end
      item
        Name = 'dt_pagamento'
        DataType = ftDateTime
      end
      item
        Name = 'dt_estorno'
        DataType = ftDateTime
      end
      item
        Name = 'dt_prevista'
        DataType = ftDateTime
      end
      item
        Name = 'ref_pedido'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'entrega'
        DataType = ftLargeint
      end
      item
        Name = 'tipo'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'valor'
        DataType = ftFloat
      end
      item
        Name = 'dt_liberacao'
        DataType = ftDateTime
      end>
    Left = 384
    Top = 200
    object qImportaPlandt_pedido: TDateTimeField
      FieldName = 'dt_pedido'
    end
    object qImportaPlandt_pagamento: TDateTimeField
      FieldName = 'dt_pagamento'
    end
    object qImportaPlandt_estorno: TDateTimeField
      FieldName = 'dt_estorno'
    end
    object qImportaPlandt_prevista: TDateTimeField
      FieldName = 'dt_prevista'
    end
    object qImportaPlanref_pedido: TStringField
      FieldName = 'ref_pedido'
    end
    object qImportaPlanentrega: TLargeintField
      FieldName = 'entrega'
    end
    object qImportaPlantipo: TStringField
      FieldName = 'tipo'
    end
    object qImportaPlanvalor: TFloatField
      FieldName = 'valor'
    end
    object qImportaPlandt_liberacao: TDateTimeField
      FieldName = 'dt_liberacao'
    end
  end
end
