unit Main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, ComCtrls, xDBTree, Menus, DlgTree, ImgList;

type
  TMainForm = class(TForm)
    tbTree: TTable;
    dsTree: TDataSource;
    tv: TXDBTreeView;
    ImageList: TImageList;
    pmTree: TPopupMenu;
    pmtvAdd: TMenuItem;
    pmtvEdit: TMenuItem;
    pmtvRemove: TMenuItem;
    pmTree1: TMenuItem;
    pmUp: TMenuItem;
    pmDown: TMenuItem;
    tb: TTable;
    tbID: TIntegerField;
    tbOwnerID: TIntegerField;
    tbPictureID: TIntegerField;
    tbDescription: TStringField;
    tbScript: TMemoField;
    procedure FormCreate(Sender: TObject);
    procedure pmtvAddClick(Sender: TObject);
    procedure pmtvEditClick(Sender: TObject);
    procedure pmtvRemoveClick(Sender: TObject);
    procedure pmUpClick(Sender: TObject);
    procedure pmDownClick(Sender: TObject);
    procedure tvDblClick(Sender: TObject);
  private
  public
    AppPath: String;
    procedure CreateTable(tbName: String);
    procedure MoveNode(iStep: Integer);
    procedure ExChangeNode(ID1,ID2: Integer);
  end;

var
  MainForm: TMainForm;

function GetMaxID(TableName: String): LongInt;
  
implementation
{$R *.DFM}

function GetMaxID(TableName: String): LongInt;
begin
  with TQuery.Create(Application) do
  try
    Screen.Cursor:=crHourGlass;
    SQL.Add('SELECT MAX(ID) FROM "' + TableName + '"');
    Open;
    result := Fields[0].AsInteger + 1;
  finally
    Close;
    Free;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TMainForm.ExChangeNode(ID1,ID2: Integer);
Begin
  IF ID2=ID1 Then
    Exit;
  try
    tv.Items.BeginUpdate;
    tb.Filtered:=True;
    tb.Filter:='OwnerID='+IntToStr(ID1);
    tb.First;
    While not tb.EOF Do
    begin
      tb.Edit;
      tbOwnerID.ASInteger:=-1;
      tb.Post;
    end;

    tb.Filter:='OwnerID='+IntToStr(ID2);
    tb.First;
    While not tb.EOF Do
    begin
      tb.Edit;
      tbOwnerID.ASInteger:=ID1;
      tb.Post;
    end;

    tb.Filter:='OwnerID=-1';
    tb.First;
    While not tb.EOF Do
    begin
      tb.Edit;
      tbOwnerID.ASInteger:=ID2;
      tb.Post;
    end;
    tb.Filtered:=False;

    tb.FindKey([ID1]);
    tb.Edit;
    tbID.AsInteger:=-1;
    tb.Post;

    tb.FindKey([ID2]);
    tb.Edit;
    tbID.AsInteger:=ID1;
    tb.Post;

    tb.FindKey([-1]);
    tb.Edit;
    tbID.AsInteger:=ID2;
    tb.Post;

    tv.RefreshTree;
    tv.Selected:=tv.GetNodeAtID(ID2);
  finally
    tb.Filtered:=False;
    tv.Items.EndUpdate;
  end;
End;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  AppPath:=ExtractFilePath(Application.ExeName);
  IF not FileExists(AppPath+'DBTree.db') Then
    CreateTable(AppPath+'DBTree.db');
  tbTree.Open;
  tb.Open;
  tv.FullExpand;
end;

procedure TMainForm.MoveNode(iStep: Integer);
var
  nd: TTreeNode;
  i,j: Integer;
Begin
  j:=0;
  nd:=tv.Selected;
  IF nd.Parent<>Nil Then
  begin
    nd:=nd.Parent;
    For i:=0 To nd.Count-1 Do
      IF nd.Item[i]=tv.Selected Then
      begin
        j:=i+iStep;
//        j:=Min(nd.Count-1,j);
//        j:=Max(0,j);
      end;
  end;
  ExChangeNode((tv.Selected as TXTreeNode).ID, (nd.Item[j] as TXTreeNode).ID);
End;

procedure TMainForm.CreateTable(tbName: String);
Begin
  With TQuery.Create(Application) Do
  try
    SQL.Add('CREATE TABLE "'+ tbName +'" (');
    SQL.Add('ID Integer,');
    SQL.Add('OwnerID Integer,');
    SQL.Add('PictureID Integer,');
    SQL.Add('Description char(60),');
    SQL.Add('Script Blob(10,1),');
    SQL.Add('primary key (id)');
    SQL.Add(')');
    ExecSQL;
  finally
    Free;
  end;
End;


procedure TMainForm.pmtvAddClick(Sender: TObject);
var
  AnID: Integer;
begin
  IF tv.Selected=Nil Then
    AnID:=-1
  Else
    AnID:=TXTreeNode(tv.Selected).ID;
  if ExecTreeDlg(-1, AnID) then
    tv.RefreshTree;
end;

procedure TMainForm.pmtvEditClick(Sender: TObject);
begin
  if ExecTreeDlg(TXTreeNode(tv.Selected).ID, TXTreeNode(tv.Selected).OwnerID) then
    tv.RefreshTree;
end;

procedure TMainForm.pmtvRemoveClick(Sender: TObject);
begin
  if Application.MessageBox('Delete Item?', 'Confirm',
                            mb_IconQuestion + mb_YesNo) = idYes then
  with TQuery.Create(Application) do
  try
    SQL.Add('DELETE FROM "DBTree" WHERE ID = :ID');
    Params[0].AsInteger := TXTreeNode(tv.Selected).ID;
    ExecSQL;
  finally
    Free;
    tv.RefreshTree;
  end;
end;

procedure TMainForm.pmUpClick(Sender: TObject);
begin
  MoveNode(-1);
end;

procedure TMainForm.pmDownClick(Sender: TObject);
begin
  MoveNode(1);
end;

procedure TMainForm.tvDblClick(Sender: TObject);
begin
  pmtvEditClick(Nil);
end;

end.
