//----------------------------------------------------------------------------------------
//  "DB Дерево своими руками"
//
// Специально для Королевства Delphi http://www.delphikingdom.com
//
// Елена Филиппова , 27 марта 2001 г.
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
// Полная перестройка дерева
//----------------------------------------------------------------------------------------
Procedure TFormTree.RebuildTreeCompanies;
Begin

    TreeCompanies.Items.Clear;

    // Принудительное раскрытие самого верхнего уровня
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
       // ID родительской ветки , для нее и ищем все дочерние
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

    // Для самого верхнего уровня выбрать только тех,
    // кто не имеет родителей.
    IF Node = nil Then ID:=0
    Else ID:=Integer(Node.Data);

    qTreeCompanies.Close;
    qTreeCompanies.ParamByName('ParentID').AsInteger:=ID;
    qTreeCompanies.Open;

    // Для каждой строки из полученного набора данных
    // формируем ветвь в TreeView, как дочерние ветки к той,
    // которую мы только что "раскрыли"

    TreeCompanies.Items.BeginUpdate;

    For i:=1 To qTreeCompanies.RecordCount Do
    Begin

       // Запишем в поле Data ветки ее идентификационный номер(ID) в таблице
       TreeNode:=TreeCompanies.Items.AddChildObject(Node ,
                                  qTreeCompanies.FieldByName('Name').AsString ,
                                  Pointer(qTreeCompanies.FieldByName('ID').AsInteger));

       TreeNode.ImageIndex:=1;
       TreeNode.SelectedIndex:=2;

       // Добавим фиктивную (пустую) дочернюю ветвь только для того,
       // чтобы был отрисован [+] на ветке и ее можно было бы раскрыть
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

  // Если первая дочерняя ветвь той, которую мы хотим раскрыть, пустая
  // значит это фиктивная ветвь, дерево еще не достраивалось по этой ветви.
  // Удаляем фиктивную ветвь и "разворачиваем" дерево вглубь той ветки,
  // на которой стоим
  IF Node.getFirstChild.Data = nil
  Then Begin
          Node.DeleteChildren;
          ExpandLevel(Node);
       End;

end;
//----------------------------------------------------------------------------------------
//  Шаг внуть дерева - имитация проваливания на более низкую ветвь
//----------------------------------------------------------------------------------------
procedure TFormTree.gridCompaniesDblClick(Sender: TObject);
Var ID , i : Integer;
    Allow  : Boolean;
begin
          ID:=qCompanies.FieldByName('ID').AsInteger;

          // принудительное "невидимое" раскрытие той ветки, на которой стоим
          TreeCompanies.OnExpanding(TreeCompanies ,TreeCompanies.Selected , Allow);

          // Перебираем все получившиеся дочерние ветки и ищем ту, ID которой
          // совпадает с ID строки в правой таблице. То есть ищем ветку в дереве,
          // которая соответсвует той записи в таблице, на которой мы стоим
          // Как только нашли, визуально раскрываем ветку и делаем ее выделенной,
          // то есть визуально "встаем" на нее в дереве
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

   // Считываем  настройку аналитики
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

    // Фиктивный верхний уровень, который существует всегда и содержит все документы
    // То есть, нет фиксации никакой аналитики.
    TreeNode:=TreeAnalytic.Items.AddChildObject( nil , 'Все документы' , nil );
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

     Level:=Node.Level + 1; // уровень, который будет раскрываться
     // Самому первому аналитическому признаку в списке  ListEntities
     // соответсвует _второй_ физический уровень веток дерева.
     // Так как самый верхний уровень дерева фиктивный -"все документы"
     // Отсюда и игра с (+/-) 1 при обращении к списку

     qTreeAnalytic.Close;

     // Определим, на каком типе уровня мы сейчас находимся
     //
     IF Level > ListEntities.Count
     Then Begin  // Уровень документов, аналитка закончилась
            Sql:='SELECT * FROM Documents Where '+ GetSqlPath(Node);
            Name:= 'DocumentID';
            ImageIndex:=3;
          End
     Else Begin // Очередной уровень аналитики
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

     // Получен очередной уровень ветвей дерева
     For i:=1 To qTreeAnalytic.RecordCount Do
     Begin
         NewItem:=List.AddItem(qTreeAnalytic.FieldByName(Name).AsInteger , Name);

         TreeNode:=TreeAnalytic.Items.AddChildObject( Node ,
                              qTreeAnalytic.FieldByName('Name').AsString, NewItem );
         TreeNode.ImageIndex:=ImageIndex;
         TreeNode.SelectedIndex:=TreeNode.ImageIndex;

         // Фиктивная дочерняя ветка ТОЛЬКО для уровней аналитики,
         // так как документы - последний уровень, за которым ничего и не может быть
         IF Level <= ListEntities.Count
         Then TreeAnalytic.Items.AddChild(TreeNode , '' );

         qTreeAnalytic.Next;
     End;

     TreeAnalytic.Items.EndUpdate;

End;
//----------------------------------------------------------------------------------------
procedure TFormTree.TreeAnalyticChange(Sender: TObject; Node: TTreeNode);
begin

    // Получаем полный путь от корня до той ветке, на которой стоим
    // и переоткрываем запрос со списком документов, которые имеют
    // отношение ТОЛЬКО к этому пути
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
// Получаем полный путь от корня до указанной ветки дерева
// Полный путь это есть зафиксированные значения для каждого уровня
// аналитики.
// Эти значения необходимы для того, чтобы верно построить запрос. То есть мы
// фактически формируем дополнительный фильтр для последующих выборок
//----------------------------------------------------------------------------------------
Function TFormTree.GetSqlPath( Node : TTreeNode ) : String;
Begin
   Result:=' 0=0 ' ;

   // Участвуют все ветви дерева, кроме самого верхнего фиктивного уровня
   While Node.Level > 0   Do
   Begin
      Result:= Result + ' AND ' +
               'Documents.' + TListsItem(Node.Data).Name + '=' +
                              TListsItem(Node.Data).AsString ;

      // Делаем шаг назад по ветке дерева
      Node:=Node.Parent;
   End;

End;
//----------------------------------------------------------------------------------------
procedure TFormTree.TreeAnalyticExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
begin

  // Если первая дочерняя ветвь той, которую мы хотим раскрыть, пустая
  // значит это фиктивная ветвь, дерево еще не достраивалось по этой ветви.
  // Удаляем фиктивную ветвь и "разворачиваем" дерево вглубь той ветки,
  // на которой стоим
  IF Node.getFirstChild.Data = nil
  Then Begin
          Node.DeleteChildren;
          ExpandLevelAnalytic(Node);
       End;

end;
//----------------------------------------------------------------------------------------
end.
