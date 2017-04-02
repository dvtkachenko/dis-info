//----------------------------------------------------------------------------------------
//  "DB ������ ������ ������"
//
// ���������� ��� ����������� Delphi http://www.delphikingdom.com
//
// ����� ��������� , 27 ����� 2001 �.
//
//----------------------------------------------------------------------------------------
unit TreeUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, ComCtrls, ToolWin, DBTables, Grids, DBGrids, ExtCtrls, ImgList, Menus,
  RxQuery , ListUtils, setupEntities;

type
  TFormTree = class(TForm)
    PageControl: TPageControl;
    Panel1: TPanel;
    tshCompany: TTabSheet;
    tshAnalytic: TTabSheet;
    TreeCompanies: TTreeView;
    Splitter1: TSplitter;
    gridCompanies: TDBGrid;
    qCompanies: TQuery;
    DataSource1: TDataSource;
    qTreeCompanies: TQuery;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    qCompaniesID: TIntegerField;
    qCompaniesName: TStringField;
    qCompaniesParentID: TIntegerField;
    ImageList1: TImageList;
    PopupMenu1: TPopupMenu;
    nEdit: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    TreeAnalytic: TTreeView;
    Splitter2: TSplitter;
    DBGrid1: TDBGrid;
    qTreeAnalytic: TQuery;
    DataSource2: TDataSource;
    qDocument: TRxQuery;
    qDocumentDocumentID: TIntegerField;
    qDocumentName: TStringField;
    qDocumentCityID: TIntegerField;
    qDocumentClientID: TIntegerField;
    qDocumentGoodsID: TIntegerField;
    ToolButton3: TToolButton;
    procedure ToolButton1Click(Sender: TObject);
    procedure TreeCompaniesChange(Sender: TObject; Node: TTreeNode);
    procedure TreeCompaniesExpanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: Boolean);
    procedure gridCompaniesDblClick(Sender: TObject);
    procedure nEditClick(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure TreeAnalyticChange(Sender: TObject; Node: TTreeNode);
    procedure qDocumentBeforeOpen(DataSet: TDataSet);
    procedure TreeAnalyticExpanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: Boolean);
  private
    ListEntities : TEntityLists;
    List         : TLists;

    MacroWhere   : String;

    Procedure RebuildTreeCompanies;
    Procedure ExpandLevel( Node : TTreeNode);

    Procedure RebuildTreeAnalytic;
    Procedure ExpandLevelAnalytic(Node : TTreeNode );
    Function GetSqlPath( Node : TTreeNode ) : String;

  public
    { Public declarations }
  end;

var
  FormTree: TFormTree;

implementation
{$R *.DFM}
Uses  EditUnit ;
//----------------------------------------------------------------------------------------
// ������ ����������� ������
//----------------------------------------------------------------------------------------
Procedure TFormTree.RebuildTreeCompanies;
Begin

    TreeCompanies.Items.Clear;

    // �������������� ��������� ������ �������� ������
    ExpandLevel(nil);
    TreeCompanies.Selected:=TreeCompanies.Items[0];

End;
//----------------------------------------------------------------------------------------
procedure TFormTree.ToolButton1Click(Sender: TObject);
begin

   IF PageControl.ActivePage = tshCompany
   Then RebuildTreeCompanies
   Else RebuildTreeAnalytic;

end;
//----------------------------------------------------------------------------------------
procedure TFormTree.TreeCompaniesChange(Sender: TObject; Node: TTreeNode);
Var ID : Integer;
begin

   IF TTreeView(Sender).Selected <> nil Then
   Begin
       // ID ������������ ����� , ��� ��� � ���� ��� ��������
       ID:=Integer(TTreeView(Sender).Selected.Data);

       qCompanies.Close;
       qCompanies.ParamByName('ParentID').AsInteger:=ID;
       qCompanies.Open;

   End;

end;
//----------------------------------------------------------------------------------------
Procedure TFormTree.ExpandLevel( Node : TTreeNode);
Var ID , i   : Integer;
    TreeNode : TTreeNode;
Begin

    // ��� ������ �������� ������ ������� ������ ���,
    // ��� �� ����� ���������.
    IF Node = nil Then ID:=0
    Else ID:=Integer(Node.Data);

    qTreeCompanies.Close;
    qTreeCompanies.ParamByName('ParentID').AsInteger:=ID;
    qTreeCompanies.Open;

    // ��� ������ ������ �� ����������� ������ ������
    // ��������� ����� � TreeView, ��� �������� ����� � ���,
    // ������� �� ������ ��� "��������"

    TreeCompanies.Items.BeginUpdate;

    For i:=1 To qTreeCompanies.RecordCount Do
    Begin

       // ������� � ���� Data ����� �� ����������������� �����(ID) � �������
       TreeNode:=TreeCompanies.Items.AddChildObject(Node ,
                                  qTreeCompanies.FieldByName('Name').AsString ,
                                  Pointer(qTreeCompanies.FieldByName('ID').AsInteger));

       TreeNode.ImageIndex:=1;
       TreeNode.SelectedIndex:=2;

       // ������� ��������� (������) �������� ����� ������ ��� ����,
       // ����� ��� ��������� [+] �� ����� � �� ����� ���� �� ��������
       TreeCompanies.Items.AddChildObject(TreeNode , '' , nil);

       qTreeCompanies.Next;
    End;

    TreeCompanies.Items.EndUpdate;

End;
//----------------------------------------------------------------------------------------
procedure TFormTree.TreeCompaniesExpanding(Sender: TObject;
  Node: TTreeNode; var AllowExpansion: Boolean);
begin
  IF Node = nil Then Exit;

  // ���� ������ �������� ����� ���, ������� �� ����� ��������, ������
  // ������ ��� ��������� �����, ������ ��� �� ������������� �� ���� �����.
  // ������� ��������� ����� � "�������������" ������ ������ ��� �����,
  // �� ������� �����
  IF Node.getFirstChild.Data = nil
  Then Begin
          Node.DeleteChildren;
          ExpandLevel(Node);
       End;

end;
//----------------------------------------------------------------------------------------
//  ��� ����� ������ - �������� ������������ �� ����� ������ �����
//----------------------------------------------------------------------------------------
procedure TFormTree.gridCompaniesDblClick(Sender: TObject);
Var ID , i : Integer;
    Allow  : Boolean;
begin
          ID:=qCompanies.FieldByName('ID').AsInteger;

          // �������������� "���������" ��������� ��� �����, �� ������� �����
          TreeCompanies.OnExpanding(TreeCompanies ,TreeCompanies.Selected , Allow);

          // ���������� ��� ������������ �������� ����� � ���� ��, ID �������
          // ��������� � ID ������ � ������ �������. �� ���� ���� ����� � ������,
          // ������� ������������ ��� ������ � �������, �� ������� �� �����
          // ��� ������ �����, ��������� ���������� ����� � ������ �� ����������,
          // �� ���� ��������� "������" �� ��� � ������
          FOR i:=0 To TreeCompanies.Selected.Count-1 Do
          IF Integer(TreeCompanies.Selected.Item[i].Data) = ID
          Then Begin
                TreeCompanies.Selected.Item[i].Expand(False);
                TreeCompanies.Selected.Item[i].Selected:=True;
                TreeCompanies.Repaint;
                Exit;
               End;


end;
//----------------------------------------------------------------------------------------
procedure TFormTree.nEditClick(Sender: TObject);
begin
   IF ModifyCompany(Integer(TreeCompanies.Selected.Data)) Then RebuildTreeCompanies;
end;
//----------------------------------------------------------------------------------------
procedure TFormTree.N3Click(Sender: TObject);
Var ID : Integer;
begin
   IF TreeCompanies.Selected <> nil
   Then ID:=Integer(TreeCompanies.Selected.Data)
   Else ID:=0;

   IF ModifyCompany( 0 , ID) Then RebuildTreeCompanies;
end;
//----------------------------------------------------------------------------------------
procedure TFormTree.PopupMenu1Popup(Sender: TObject);
begin
   nEdit.Enabled:=(TreeCompanies.Selected <> nil);
end;
//----------------------------------------------------------------------------------------
procedure TFormTree.ToolButton3Click(Sender: TObject);
begin
    IF SetEntities(ListEntities) Then RebuildTreeAnalytic;
end;
//----------------------------------------------------------------------------------------
procedure TFormTree.FormCreate(Sender: TObject);
begin
   ListEntities:=TEntitylists.Create(TEntityItem);
   List:=Tlists.Create(TListsItem);

   // ���������  ��������� ���������
   GetEntities(ListEntities);

   RebuildTreeCompanies;
   RebuildTreeAnalytic;
end;
//----------------------------------------------------------------------------------------
procedure TFormTree.FormDestroy(Sender: TObject);
begin
  ListEntities.Free;
  List.Free;
end;
//----------------------------------------------------------------------------------------
Procedure TFormTree.RebuildTreeAnalytic;
Var TreeNode : TTreeNode;
Begin

    TreeAnalytic.Items.Clear;
    List.Clear;

    // ��������� ������� �������, ������� ���������� ������ � �������� ��� ���������
    // �� ����, ��� �������� ������� ���������.
    TreeNode:=TreeAnalytic.Items.AddChildObject( nil , '��� ���������' , nil );
    TreeNode.ImageIndex:=3;
    TreeNode.SelectedIndex:=3;
    TreeAnalytic.Items.AddChildObject(TreeNode , '' , nil );

    TreeAnalytic.Selected:=TreeAnalytic.Items[0];

End;
//----------------------------------------------------------------------------------------
Procedure TFormTree.ExpandLevelAnalytic(Node : TTreeNode );
Var NewItem    : TListsItem;
    ImageIndex ,
    Level , i  : Integer;
    TreeNode   : TTreeNode;
    Sql,Name   : String;
Begin

     IF Node = nil Then Exit;

     TreeAnalytic.Items.BeginUpdate;

     Level:=Node.Level + 1; // �������, ������� ����� ������������
     // ������ ������� �������������� �������� � ������  ListEntities
     // ������������ _������_ ���������� ������� ����� ������.
     // ��� ��� ����� ������� ������� ������ ��������� -"��� ���������"
     // ������ � ���� � (+/-) 1 ��� ��������� � ������

     qTreeAnalytic.Close;

     // ���������, �� ����� ���� ������ �� ������ ���������
     //
     IF Level > ListEntities.Count
     Then Begin  // ������� ����������, �������� �����������
            Sql:='SELECT * FROM Documents Where '+ GetSqlPath(Node);
            Name:= 'DocumentID';
            ImageIndex:=3;
          End
     Else Begin // ��������� ������� ���������
            Sql:='SELECT DISTINCT '+ ListEntities[Level-1].AsString + '.* ' +
                 ' FROM Documents , ' +  ListEntities[Level-1].AsString + ' WHERE ' +
                 ListEntities[Level-1].AsString + '.' +ListEntities[Level-1].Name + '=' +
                 'Documents.'+ListEntities[Level-1].Name + ' AND ' + GetSqlPath(Node) ;
            Name:=ListEntities[Level-1].Name;
            ImageIndex:=ListEntities[Level-1].ImageIndex;
          End;

     qTreeAnalytic.Sql.Clear;
     qTreeAnalytic.Sql.Add(Sql);

     qTreeAnalytic.Open;

     // ������� ��������� ������� ������ ������
     For i:=1 To qTreeAnalytic.RecordCount Do
     Begin
         NewItem:=List.AddItem(qTreeAnalytic.FieldByName(Name).AsInteger , Name);

         TreeNode:=TreeAnalytic.Items.AddChildObject( Node ,
                              qTreeAnalytic.FieldByName('Name').AsString, NewItem );
         TreeNode.ImageIndex:=ImageIndex;
         TreeNode.SelectedIndex:=TreeNode.ImageIndex;

         // ��������� �������� ����� ������ ��� ������� ���������,
         // ��� ��� ��������� - ��������� �������, �� ������� ������ � �� ����� ����
         IF Level <= ListEntities.Count
         Then TreeAnalytic.Items.AddChild(TreeNode , '' );

         qTreeAnalytic.Next;
     End;

     TreeAnalytic.Items.EndUpdate;

End;
//----------------------------------------------------------------------------------------
procedure TFormTree.TreeAnalyticChange(Sender: TObject; Node: TTreeNode);
begin

    // �������� ������ ���� �� ����� �� ��� �����, �� ������� �����
    // � ������������� ������ �� ������� ����������, ������� �����
    // ��������� ������ � ����� ����
    MacroWhere:=GetSqlPath(Node);
    qDocument.Close;
    qDocument.Open;

end;
//----------------------------------------------------------------------------------------
procedure TFormTree.qDocumentBeforeOpen(DataSet: TDataSet);
begin

    TRXQuery(DataSet).MacroByName('MacroWhere').AsString:=MacroWhere;

end;
//----------------------------------------------------------------------------------------
// �������� ������ ���� �� ����� �� ��������� ����� ������
// ������ ���� ��� ���� ��������������� �������� ��� ������� ������
// ���������.
// ��� �������� ���������� ��� ����, ����� ����� ��������� ������. �� ���� ��
// ���������� ��������� �������������� ������ ��� ����������� �������
//----------------------------------------------------------------------------------------
Function TFormTree.GetSqlPath( Node : TTreeNode ) : String;
Begin
   Result:=' 0=0 ' ;

   // ��������� ��� ����� ������, ����� ������ �������� ���������� ������
   While Node.Level > 0   Do
   Begin
      Result:= Result + ' AND ' +
               'Documents.' + TListsItem(Node.Data).Name + '=' +
                              TListsItem(Node.Data).AsString ;

      // ������ ��� ����� �� ����� ������
      Node:=Node.Parent;
   End;

End;
//----------------------------------------------------------------------------------------
procedure TFormTree.TreeAnalyticExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
begin

  // ���� ������ �������� ����� ���, ������� �� ����� ��������, ������
  // ������ ��� ��������� �����, ������ ��� �� ������������� �� ���� �����.
  // ������� ��������� ����� � "�������������" ������ ������ ��� �����,
  // �� ������� �����
  IF Node.getFirstChild.Data = nil
  Then Begin
          Node.DeleteChildren;
          ExpandLevelAnalytic(Node);
       End;

end;
//----------------------------------------------------------------------------------------
end.
