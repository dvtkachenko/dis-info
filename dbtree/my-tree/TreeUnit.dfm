object FormTree: TFormTree
  Left = 284
  Top = 199
  Width = 574
  Height = 366
  Caption = 'DB-������ ������ ������'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  ShowHint = True
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl: TPageControl
    Left = 0
    Top = 36
    Width = 566
    Height = 303
    ActivePage = tshCompany
    Align = alClient
    TabOrder = 0
    object tshCompany: TTabSheet
      Caption = '������ �������������'
      object Splitter1: TSplitter
        Left = 246
        Top = 0
        Width = 3
        Height = 275
        Cursor = crHSplit
      end
      object TreeCompanies: TTreeView
        Left = 0
        Top = 0
        Width = 246
        Height = 275
        Align = alLeft
        Images = ImageList1
        Indent = 19
        PopupMenu = PopupMenu1
        StateImages = ImageList1
        TabOrder = 0
        OnChange = TreeCompaniesChange
        OnExpanding = TreeCompaniesExpanding
      end
      object gridCompanies: TDBGrid
        Left = 249
        Top = 0
        Width = 309
        Height = 275
        Align = alClient
        DataSource = DataSource1
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        OnDblClick = gridCompaniesDblClick
        Columns = <
          item
            Expanded = False
            FieldName = 'ID'
            Title.Caption = '�'
            Width = 20
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Name'
            Title.Caption = '�������� �������������'
            Width = 200
            Visible = True
          end>
      end
    end
    object tshAnalytic: TTabSheet
      Caption = '������ ������������� ���������'
      ImageIndex = 1
      object Splitter2: TSplitter
        Left = 236
        Top = 0
        Width = 3
        Height = 275
        Cursor = crHSplit
      end
      object TreeAnalytic: TTreeView
        Left = 0
        Top = 0
        Width = 236
        Height = 275
        Align = alLeft
        Images = ImageList1
        Indent = 19
        TabOrder = 0
        OnChange = TreeAnalyticChange
        OnExpanding = TreeAnalyticExpanding
      end
      object DBGrid1: TDBGrid
        Left = 239
        Top = 0
        Width = 319
        Height = 275
        Align = alClient
        DataSource = DataSource2
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = 'DocumentID'
            Title.Caption = '�'
            Width = 29
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Name'
            Title.Caption = '�������� ���������'
            Width = 200
            Visible = True
          end>
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 566
    Height = 36
    Align = alTop
    TabOrder = 1
    object ToolBar1: TToolBar
      Left = 1
      Top = 1
      Width = 564
      Height = 29
      Caption = 'ToolBar1'
      EdgeBorders = [ebBottom]
      Images = ImageList1
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      object ToolButton2: TToolButton
        Left = 0
        Top = 2
        Width = 8
        Caption = 'ToolButton2'
        ImageIndex = 1
        Style = tbsSeparator
      end
      object ToolButton1: TToolButton
        Left = 8
        Top = 2
        Hint = '�����������  ������'
        Caption = 'ToolButton1'
        ImageIndex = 0
        OnClick = ToolButton1Click
      end
      object ToolButton3: TToolButton
        Left = 31
        Top = 2
        Hint = '��������� ���������'
        Caption = 'ToolButton3'
        ImageIndex = 7
        OnClick = ToolButton3Click
      end
    end
  end
  object qCompanies: TQuery
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'Select * from COMPANY'
      'Where ParentID=:ParentID')
    Left = 265
    Top = 296
    ParamData = <
      item
        DataType = ftInteger
        Name = 'ParentID'
        ParamType = ptInput
        Value = 0
      end>
    object qCompaniesID: TIntegerField
      FieldName = 'ID'
      Origin = 'TREEDB."company.DB".ID'
    end
    object qCompaniesName: TStringField
      FieldName = 'Name'
      Origin = 'TREEDB."company.DB".Name'
      Size = 255
    end
    object qCompaniesParentID: TIntegerField
      FieldName = 'ParentID'
      Origin = 'TREEDB."company.DB".ParentID'
    end
  end
  object DataSource1: TDataSource
    DataSet = qCompanies
    Left = 294
    Top = 296
  end
  object qTreeCompanies: TQuery
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'Select *'
      'From COMPANY'
      'Where ParentID=:ParentID')
    Left = 14
    Top = 293
    ParamData = <
      item
        DataType = ftInteger
        Name = 'ParentID'
        ParamType = ptInput
        Value = 0
      end>
  end
  object ImageList1: TImageList
    Left = 185
    Top = 294
    Bitmap = {
      494C010108000900040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000003000000001001800000000000024
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000007B7B7B0000007B7B7B7B7B7B7B7B7B000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000084848400000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000BDBDBD7B7B7B0000007B7B7B000000000000000000000000000000000000
      0000007B7B7B0000007B7B7B00000000000000000000000000000000FFFF00FF
      FF00FFFF00FFFF00FFFF00000000FFFF00FFFF00000000848484848400000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000008400000000000000000000000000000000000000000000
      00000000000000007B7B7B7B7B7B000000000000000000000000000000000000
      0000000000007B7B7B0000000000000000000000000000000000000000000000
      0000000000000000FFFF00000000FFFF00FFFF00000000848484848400000000
      0000000000000000008400008400008400008400008400008400000000008400
      0084000084000084000000000000000000000000000000000000000000000000
      FF0000FF0000FF0000840000FF0000FF0000FF00000000000000000000000000
      0000BDBDBDBDBDBD7B7B7B0000000000007B7B7B000000000000000000000000
      0000007B7B7B0000000000007B7B7B7B7B7B000000000000000000000000FFFF
      00FFFF0000000000FFFF00000000FFFF00FFFF00000000848484848400000000
      0000000000000000000000000000008400008400000000000000000000008400
      0084000000000000000000000000000000000000000000000000000000000000
      FF0000FF0000FF0000840000FF0000FF0000FF0000000000FF00000000000000
      00007B7B7B0000000000000000007B7B7B000000000000000000000000000000
      0000000000000000007B7B7B7B7B7B000000000000000000000000000000FFFF
      00FFFF0000000000FFFF00000000FFFF00FFFF00000000848484848400000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      FF0000FF0000FF0000840000FF0000FF0000FF0000000000FF00000000000000
      00007B7B7B0000000000007B7B7B0000007B7B7B000000000000000000000000
      000000000000000000BDBDBD0000007B7B7B0000000000000000000000000000
      0000000000000000FFFF00000000000000000000000000848484848400000000
      0000000000000000000000000000FFFFFF848484000000008400000000FFFFFF
      8484840000000084000084000084000000000000000000000000000000000000
      FF0000FF0000FF0000840000FF0000FF0000FF0000000000FF00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000007B7B7B7B7B7BBDBDBD000000BDBDBD0000000000000000000084840084
      8400848400848400848400848400848400848400848400848484848484848400
      0000000000000000000000000000848484FFFFFF000000008400000000848484
      FFFFFF0000000084000084000084000000000000000000000000000000000000
      FF0000FF0000FF0000840000FF0000FF0000FF0000000000FF00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000BDBDBD0000007B7B7B0000000000000000840000840000
      FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000840000FF84848400
      0000000000000000000000000000000000000000000000008400000000000000
      0000000000000084000084000000000000000000000000000000000000000000
      000000000000000000840000000000000000000000000000FF00000000000000
      00000000007B7B7B0000000000007B7B7B0000000000007B7B7B0000007B7B7B
      0000000000000000007B7B7BBDBDBD0000000000000000000000000000840000
      840000FF0000000000000000000000840000FF0000840000FF84848400000000
      0000000000000000000000000000000000000000008400008400008400000000
      0000000084000084000000000000000000000000000000000000000000000000
      000000000000FF0000FF0000840000FF0000FF0000FF0000FF00000000000000
      00007B7B7B0000000000000000007B7B7B0000000000000000007B7B7B000000
      0000007B7B7B0000000000007B7B7BBDBDBD0000000000000000000000000000
      840000840000FFFFFF000000000000FF0000840000FF84848400000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000008400848400008400008400000000000000000000000000
      00000000000000007B7B7B7B7B7B7B7B7B7B7B7B7B7B7B0000000000007B7B7B
      0000000000007B7B7B0000000000000000000000000000000000000000000000
      000000840000840000FF0000FF0000840000FF84848484848400000000000000
      0000000000000000000000000000000000000000FFFFFF848484000000000000
      000000FFFFFF8484840000000000000000000000000000000000000000000000
      0000000000008400008400008400000000008400000000000000000000000000
      00000000007B7B7B7B7B7B0000000000000000007B7B7B7B7B7B000000000000
      0000007B7B7B0000007B7B7B0000000000000000000000000000000000000000
      000000000000840000840000840000FF00000084848484848400000000000000
      0000000000000000000000000000000000000000848484FFFFFF000000000000
      000000848484FFFFFF0000000000000000000000000000000000000000000000
      0000008400008400000000008400000000008400000000000000000000000000
      0000000000BDBDBD0000007B7B7B7B7B7B7B7B7B0000007B7B7B000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000840000FF84848400000084848484848400000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000008400000000000000000000000000000000
      00007B7B7BBDBDBD000000BDBDBD0000007B7B7B0000007B7B7B7B7B7B7B7B7B
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000084848400000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000BDBDBD0000007B7B7BBDBDBD7B7B7B0000007B7B7B000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000007B7B7BBDBDBD0000000000000000007B7B7B7B7B7B000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFFFFFFFF008400FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFF008400008400FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000000000
      0000000000000000000000C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6
      C6C6C6C6C6C60000000000000000000000000000000000000000000084000084
      0000840000840000840000840000840000840000840000000000000000000000
      0000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFF000000000000000000000000000000000000000000FFFFFFFFFF
      FF008400008400008400008400008400008400FFFFFFFFFFFF00000000000000
      0000000000000000000000C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6
      C6C6C6C6C6C60000000000000000000000000000000000000000000084000084
      0000840000840000840000840000840000840000840000000000000000000000
      0000000000000000000000000000FFFFFF008484008484008484008484008484
      008484FFFFFF000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFF008400008400FFFFFFFFFFFFFFFFFF008400FFFFFF00000000000000
      0000000000000000000000C6C6C6C6C6C6C6C6C6C6C6C6FFFFFFC6C6C6C6C6C6
      C6C6C6C6C6C60000000000000000000000000000000000000000000084000084
      00008400008400FFFFFF00840000840000840000840000000000000000000000
      0000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFF000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFFFFFFFF008400FFFFFFFFFFFFFFFFFF008400FFFFFF00000000000000
      0000000000000000000000000000C6C6C6C6C6C6C6C6C6FFFFFFC6C6C6C6C6C6
      C6C6C60000000000000000000000000000000000000000000000000000000084
      00008400008400FFFFFF00840000840000840000000000000000000000000000
      0000000000000000000000000000FFFFFF008484008484008484008484008484
      008484FFFFFF000000000000000000000000000000000000000000FFFFFF0084
      00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF008400FFFFFF00000000000000
      0000000000000000000000000000000000000000FFFFFFFFFFFFFFFFFF000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000FFFFFFFFFFFFFFFFFF00000000000000000000000000000000000000
      0000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFF000000000000000000000000000000000000000000FFFFFF0084
      00FFFFFFFFFFFFFFFFFF008400FFFFFFFFFFFFFFFFFFFFFFFF00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000FFFFFF008484008484000000000000000000
      000000000000000000000000000000000000000000000000000000FFFFFF0084
      00FFFFFFFFFFFFFFFFFF008400008400FFFFFFFFFFFFFFFFFF00000000000000
      0000000000000000000000000000000000000000848484848484848484000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000084848484848484848400000000000000000000000000000000000000
      0000000000000000000000000000FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFF
      FFFFFF000000000000000000000000000000000000000000000000FFFFFFFFFF
      FF008400008400008400008400008400008400FFFFFFFFFFFF00000000000000
      0000000000000000000000000000000000000000848484848484848484000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000084848484848484848400000000000000000000000000000000000000
      0000000000000000000000000000FFFFFF008484008484000000FFFFFFFFFFFF
      000000000000000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFF008400008400FFFFFFFFFFFFFFFFFF00000000000000
      0000000000000000000000000000000000000000848484848484848484000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000084848484848484848400000000000000000000000000000000000000
      0000000000000000000000000000FFFFFFFFFFFFFFFFFF000000FFFFFF000000
      000000000000000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFF008400FFFFFF00000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000FFFFFFFFFFFFFFFFFF000000000000000000
      000000000000000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFF00000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000FFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000300000000100010000000000800100000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000FFFFFFFFFFFF81FEC003FFFFFFFF01E2
      C0038003FEFF07E0C0038003F01F03E0C003C087F01703F0C003F39FF01723C0
      C003E001F0173FC08001E001F017E3C08001E001FEF72230C003F003FC070020
      E007FCE7FE1F0020F007F843FC5F0062F807F843F95F001EFC07F843FFBF001F
      FE47FCE7FFFF001FFFFFFFFFFFFF007FFFFFFFFFFFFFFFFFC003FFFFFFFFFFFF
      C003C007C007E007C003C007C007E007C003C007C007E007C003C007C007E007
      C003E00FE00FE007C003F01FF01FE007C003FC7FFC7FE007C003F83FF83FE00F
      C003F83FF83FE01FC003F83FF83FE03FC003F83FF83FE07FC007FC7FFC7FE0FF
      C00FFFFFFFFFFFFFC01FFFFFFFFFFFFF00000000000000000000000000000000
      000000000000}
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 14
    Top = 260
    object nEdit: TMenuItem
      Caption = '�������������'
      OnClick = nEditClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object N3: TMenuItem
      Caption = '����� �������������'
      OnClick = N3Click
    end
  end
  object qTreeAnalytic: TQuery
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'Select * From Documents ')
    Left = 98
    Top = 295
  end
  object DataSource2: TDataSource
    DataSet = qDocument
    Left = 268
    Top = 241
  end
  object qDocument: TRxQuery
    BeforeOpen = qDocumentBeforeOpen
    DatabaseName = 'TreeDB'
    SQL.Strings = (
      'SELECT * FROM Documents'
      'WHERE %MacroWhere')
    Macros = <
      item
        DataType = ftString
        Name = 'MacroWhere'
        ParamType = ptInput
        Value = '0=0'
      end>
    Left = 297
    Top = 241
    object qDocumentDocumentID: TIntegerField
      FieldName = 'DocumentID'
      Origin = 'TREEDB."Documents.DB".DocumentID'
    end
    object qDocumentName: TStringField
      FieldName = 'Name'
      Origin = 'TREEDB."Documents.DB".Name'
      Size = 255
    end
    object qDocumentCityID: TIntegerField
      FieldName = 'CityID'
      Origin = 'TREEDB."Documents.DB".CityID'
    end
    object qDocumentClientID: TIntegerField
      FieldName = 'ClientID'
      Origin = 'TREEDB."Documents.DB".ClientID'
    end
    object qDocumentGoodsID: TIntegerField
      FieldName = 'GoodsID'
      Origin = 'TREEDB."Documents.DB".GoodsID'
    end
  end
end