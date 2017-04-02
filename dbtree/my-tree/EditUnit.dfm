object FormEdit: TFormEdit
  Left = 414
  Top = 382
  BorderStyle = bsDialog
  Caption = 'Редактирование '
  ClientHeight = 111
  ClientWidth = 394
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 9
    Top = 45
    Width = 50
    Height = 13
    Caption = 'Название'
  end
  object Label2: TLabel
    Left = 8
    Top = 79
    Width = 135
    Height = 13
    Caption = 'Головное подразделеение'
  end
  object edName: TEdit
    Left = 73
    Top = 42
    Width = 318
    Height = 21
    TabOrder = 0
    Text = 'edName'
  end
  object listParent: TDBLookupComboBox
    Left = 153
    Top = 74
    Width = 238
    Height = 21
    KeyField = 'ID'
    ListField = 'Name'
    ListSource = DataSource1
    TabOrder = 1
  end
  object BitBtn2: TBitBtn
    Left = 235
    Top = 5
    Width = 75
    Height = 25
    Caption = 'Сохранить'
    TabOrder = 2
    OnClick = BitBtn2Click
  end
  object BitBtn1: TBitBtn
    Left = 314
    Top = 5
    Width = 75
    Height = 25
    Caption = 'Отменить'
    ModalResult = 2
    TabOrder = 3
  end
  object DataSource1: TDataSource
    DataSet = cdsList
    Left = 238
    Top = 71
  end
  object qList: TQuery
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'SELECT Name , ID  from COMPANY'
      'ORDER BY Name')
    Left = 267
    Top = 70
    object qListName: TStringField
      FieldName = 'Name'
      Origin = 'TREEDB."COMPANY.DB".Name'
      Size = 255
    end
    object qListID: TIntegerField
      FieldName = 'ID'
      Origin = 'TREEDB."COMPANY.DB".ID'
    end
  end
  object qCommand: TQuery
    DatabaseName = 'TreeDB'
    Left = 9
    Top = 4
  end
  object cdsList: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'DataSetProvider1'
    AfterOpen = cdsListAfterOpen
    Left = 181
    Top = 71
    object cdsListName: TStringField
      FieldName = 'Name'
      Size = 255
    end
    object cdsListID: TIntegerField
      FieldName = 'ID'
    end
  end
  object DataSetProvider1: TDataSetProvider
    DataSet = qList
    Constraints = True
    Left = 209
    Top = 71
  end
end
