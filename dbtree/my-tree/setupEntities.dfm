object fSetUpEntities: TfSetUpEntities
  Left = 335
  Top = 330
  BorderStyle = bsDialog
  Caption = 'Настройка аналитики'
  ClientHeight = 184
  ClientWidth = 343
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object DBGrid1: TDBGrid
    Left = 0
    Top = 33
    Width = 343
    Height = 151
    Align = alClient
    DataSource = DataSource1
    PopupMenu = PopupMenu1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDrawColumnCell = DBGrid1DrawColumnCell
    OnDblClick = DBGrid1DblClick
    Columns = <
      item
        Expanded = False
        FieldName = 'IsSelect'
        Title.Caption = 'Выбрано'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Name'
        Title.Caption = 'Название аналитики'
        Width = 226
        Visible = True
      end>
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 343
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object BitBtn2: TBitBtn
      Left = 190
      Top = 4
      Width = 75
      Height = 25
      Caption = 'Сохранить'
      ModalResult = 1
      TabOrder = 0
      OnClick = BitBtn2Click
    end
    object BitBtn1: TBitBtn
      Left = 266
      Top = 4
      Width = 75
      Height = 25
      Caption = 'Отменить'
      ModalResult = 2
      TabOrder = 1
    end
  end
  object qEntities: TQuery
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'Select * from Entities Order By OrderNO')
    Left = 16
    Top = 140
  end
  object cdsEntities: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'DataSetProvider1'
    Left = 44
    Top = 140
    object cdsEntitiesEntityID: TIntegerField
      FieldName = 'EntityID'
    end
    object cdsEntitiesName: TStringField
      FieldName = 'Name'
      Size = 255
    end
    object cdsEntitiesTableName: TStringField
      FieldName = 'TableName'
      Size = 50
    end
    object cdsEntitiesKeyColumn: TStringField
      FieldName = 'KeyColumn'
      Size = 50
    end
    object cdsEntitiesIsSelect: TSmallintField
      FieldName = 'IsSelect'
      OnGetText = cdsEntitiesIsSelectGetText
    end
    object cdsEntitiesOrderNo: TIntegerField
      FieldName = 'OrderNo'
    end
    object cdsEntitiesImageIndex: TIntegerField
      FieldName = 'ImageIndex'
    end
  end
  object DataSetProvider1: TDataSetProvider
    DataSet = qEntities
    Constraints = True
    Left = 73
    Top = 140
  end
  object DataSource1: TDataSource
    DataSet = cdsEntities
    Left = 102
    Top = 140
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 16
    Top = 112
    object nUp: TMenuItem
      Tag = -1
      Caption = 'Вверх'
      OnClick = nUpClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object nDown: TMenuItem
      Tag = 1
      Caption = 'Вниз'
      OnClick = nUpClick
    end
  end
  object qCommand: TQuery
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'UPDATE Entities Set IsSelect = :IsSelect ,'
      'OrderNo =:OrderNo '
      'WHERE EntityID=:EntityID')
    Left = 130
    Top = 140
    ParamData = <
      item
        DataType = ftInteger
        Name = 'IsSelect'
        ParamType = ptInput
        Value = 0
      end
      item
        DataType = ftInteger
        Name = 'OrderNo'
        ParamType = ptInput
        Value = 0
      end
      item
        DataType = ftInteger
        Name = 'EntityID'
        ParamType = ptInput
        Value = 0
      end>
  end
end
