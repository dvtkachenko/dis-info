{*******************************************************}
{                                                       }
{         Delphi VCL eXtensions   (X)                   }
{                                                       }
{         Copyright (c) 1997 Alexander Buzaev           }
{         E-Mail: buzaev@usa.net                        }
{                                                       }
{*******************************************************}

unit xDBTree;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  CommCtrl, ComCtrls, DB, DBTables, TypInfo, DsgnIntf, Math;

type
  TNodeEvent = procedure (Sender: TObject; Node: TTreeNode) of object;
  TXTreeNode = Class;
  TXDBTreeDataLink = class;
  TXCustomDBTreeView = class(TTreeView)
  private
    FUpdateLock: Boolean;
    FDataFields: array[0..3] of String;
    FDataBaseName: String;
    FUpdateTable: String;
    FOnCreateNode: TNodeEvent;
    FOnFillNode: TNodeEvent;
    FOnDestroyNode: TNodeEvent;
    FOnBuildTree: TNotifyEvent;
    procedure SetNode(Index: integer; ANode: TXTreeNode);
    function GetNode(Index: integer): TXTreeNode;
    function GetDataSource: TDataSource;
    procedure SetDataSource(Value: TDataSource);
    procedure SetDataField(Index: Integer; Value: String);
    function GetField(Index: Integer): TField;
  protected
    FDataLink: TXDBTreeDataLink;
    FNodes: TList;
    iSaveAbsoluteIndex: Integer;
    procedure DoCreateNode(Node: TTreeNode); dynamic;
    procedure DoDestroyNode(Node: TTreeNode); dynamic;
    procedure DoFillNode(Node: TTreeNode); dynamic;
    procedure DoBuildTree; dynamic;
    procedure BuildNode(Node: TXTreeNode); virtual;
    function  CreateNode: TTreeNode; override;
    procedure CreateWnd; override;
    procedure DestroyWnd; override;
    procedure SaveState; virtual;
    procedure LoadState; virtual;
    procedure SaveTree; virtual;
    procedure ClearNodes;
    procedure CheckUpdateParams; virtual;

    property OnCreateNode: TNodeEvent read FOnCreateNode write FOnCreateNode;
    property OnFillNode: TNodeEvent read FOnFillNode write FOnFillNode;
    property OnBuildTree: TNotifyEvent read FOnBuildTree write FOnBuildTree;
    property OnDestroyNode: TNodeEvent read FOnDestroyNode write FOnDestroyNode;

    property Nodes[Index : Integer]: TXTreeNode read GetNode write SetNode; default;
    property DataSource: TDataSource read GetDataSource write SetDataSource;
    property DataFieldID: String index 0 read FDataFields[0] write SetDataField;
    property DataFieldOwnerID: String index 1 read FDataFields[1] write SetDataField;
    property DataFieldText: String index 2 read FDataFields[2] write SetDataField;
    property DataFieldImageIndex: String index 3 read FDataFields[3] write SetDataField;
    property DataBaseName: String read FDataBaseName write FDataBaseName;
    property UpdateTable: String read FUpdateTable write FUpdateTable;
    property FieldID: TField index 0 read GetField;
    property FieldOwnerID: TField index 1 read GetField;
    property FieldText: TField index 2 read GetField;
    property FieldImageIndex: TField index 3 read GetField;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure RefreshTree; virtual;
    procedure Loaded; override;
    function  GetLastNodeAtLevel(Node: TTreeNode; Level: Integer): TTreeNode; virtual;
    function  GetNodeAtID(ID: Integer): TTreeNode; virtual;
    procedure MoveOnNode(SrcID, DestID: Integer); virtual;
    procedure ExChangeNode(SrcID, DestID: Integer); virtual;
    function  GetNextNode(Node: TTreeNode; iStep: Integer): TTreeNode;
    procedure ExecSQL(aSQL: String);
    function  CreateSQL(aSQL: String): TQuery;
  published
  end;

  //¬ св€зи стем что в методе TTreeNode.Destoy ошибка (обращаютс€ к owner не провер€€€ на Nil)
  TxNodeInfo=record
    ID, OwnerID: Integer;
    ImageIndex: Integer;
    Data: Pointer;
    Text: String;
  end;

  TxTreeNode = class (TTreeNode)
  private
  public
    ID, OwnerID: Integer;
    procedure GetNodeInfo(var InfoRec: TxNodeInfo);
    constructor Create(AOwner: TTreeNodes);
  end;

  TXDBTreeView = class(TXCustomDBTreeView)
  public
    // new properties
    property FieldID;
    property FieldOwnerID;
    property FieldText;
    //inherited propertied
    property Items;
  published
    //new properties
    property DataSource;
    property DataFieldID;
    property DataFieldOwnerID;
    property DataFieldImageIndex;
    property DataFieldText;

    property OnBuildTree;
    property OnFillNode;
    //inherited propertied
    property ShowButtons;
    property BorderStyle;
    property DragCursor;
    property ShowLines;
    property ShowRoot;
    property ReadOnly;
    property DragMode;
    property HideSelection;
    property Indent;
    property OnEditing;
    property OnEdited;
    property OnExpanding;
    property OnExpanded;
    property OnCollapsing;
    property OnCompare;
    property OnCollapsed;
    property OnChanging;
    property OnChange;
    property OnDeletion;
    property OnGetImageIndex;
    property OnGetSelectedIndex;
    property Align;
    property Enabled;
    property Font;
    property Color;
    property ParentColor;
    property ParentCtl3D;
    property Ctl3D;
    property SortType;
    property TabOrder;
    property TabStop default True;
    property DataBaseName;
    property UpdateTable;
    property Visible;
    property OnClick;
    property OnEnter;
    property OnExit;
    property OnDragDrop;
    property OnDragOver;
    property OnStartDrag;
    property OnEndDrag;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnDblClick;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property PopupMenu;
    property ParentFont;
    property ParentShowHint;
    property ShowHint;
    property Images;
    property StateImages;
  end;

  TXDBTreeDataLink = class (TDataLink)
  private
    FTreeView: TXCustomDBTreeView;
  protected
    procedure ActiveChanged; override;
  end;

procedure Register;

implementation
{$R *.DCR}


