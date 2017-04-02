//----------------------------------------------------------------------------------------
//  "DB Дерево своими руками"
//
// Специально для Королевства Delphi http://www.delphikingdom.com
//
// Елена Филиппова , 27 марта 2001 г.
//
//----------------------------------------------------------------------------------------
unit setupEntities;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, Grids, DBGrids, Provider, DBClient, DBTables , ListUtils, StdCtrls,
  Buttons, Menus, ComCtrls, ToolWin, ExtCtrls;

type

  TEntityItem = class(TListsItem)
  Public
     ImageIndex : Integer;
  End;

  TEntityLists = class(TLists)
  Private
     Function GetEntityItem( Index : Integer) : TEntityItem;
  Public
     Function AddEntity( Value : Variant; AName :  String =''; ImageIndex : Integer = 0 ) : TEntityItem;
     property Items[ Index : Integer ] : TEntityItem read GetEntityItem; default;
  End;

  TfSetUpEntities = class(TForm)
    qEntities: TQuery;
    cdsEntities: TClientDataSet;
    DataSetProvider1: TDataSetProvider;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    PopupMenu1: TPopupMenu;
    nUp: TMenuItem;
    N2: TMenuItem;
    nDown: TMenuItem;
    cdsEntitiesEntityID: TIntegerField;
    cdsEntitiesName: TStringField;
    cdsEntitiesTableName: TStringField;
    cdsEntitiesKeyColumn: TStringField;
    cdsEntitiesIsSelect: TSmallintField;
    cdsEntitiesOrderNo: TIntegerField;
    cdsEntitiesImageIndex: TIntegerField;
    qCommand: TQuery;
    Panel1: TPanel;
    BitBtn2: TBitBtn;
    BitBtn1: TBitBtn;
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure cdsEntitiesIsSelectGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure nUpClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  end;

var
  fSetUpEntities: TfSetUpEntities;

Function SetEntities( var List : TEntityLists ) : Boolean;
Procedure GetEntities( var List : TEntityLists );

implementation
{$R *.DFM}
Uses DBManagerUtils;

//----------------------------------------------------------------------------------------
Procedure GetEntities( var List : TEntityLists );
Begin

    IF List = nil Then Exit;

    With GetQuery(['Select * from Entities Where IsSelect = 1 Order By OrderNO'],nil) Do
    Try

       Open;

       List.Clear;
       First;

       While NOT  EOF Do
       Begin
           List.AddEntity(FieldByName('TableName').AsString ,
                          FieldByName('KeyColumn').AsString ,
                          FieldByName('ImageIndex').AsInteger);
           Next;
       End;

    Finally

       Free;

    End;

End;
//----------------------------------------------------------------------------------------
Function SetEntities( var List : TEntityLists ) : Boolean;
Begin


    IF  fSetUpEntities = nil Then Application.CreateForm(TfSetUpEntities,fSetUpEntities);

    With fSetUpEntities Do
    Begin
       cdsEntities.Close;
       cdsEntities.Open;

       Result:=ShowModal = mrOK;
       IF Result
       Then Begin
               List.Clear;
               CdsEntities.First;

               While NOT  CdsEntities.EOF Do
               Begin
                 IF CdsEntities.FieldByName('IsSelect').AsInteger = 1
                 Then List.AddEntity(CdsEntities.FieldByName('TableName').AsString ,
                                 CdsEntities.FieldByName('KeyColumn').AsString ,
                                 CdsEntities.FieldByName('ImageIndex').AsInteger);
                 CdsEntities.Next;
               End;
            End;

    End;

End;
//----------------------------------------------------------------------------------------
procedure TfSetUpEntities.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
Var Check : Boolean;
begin

    Check:=Column.Field.DataSet.FieldByName('IsSelect').AsInteger = 1;

    IF Check
    Then TDBGrid(Sender).Canvas.Brush.Color:=clYellow;

    // если строка была выделена, оставляем "подсвеченные" цвета
    IF  gdSelected   IN State
    Then Begin
    		TDBGrid(Sender).Canvas.Brush.Color:= clHighLight;
    		TDBGrid(Sender).Canvas.Font.Color := clHighLightText;

                IF Check Then TDBGrid(Sender).Canvas.Font.Color :=clYellow;
    	End;

    TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);

end;
//----------------------------------------------------------------------------------------
procedure TfSetUpEntities.DBGrid1DblClick(Sender: TObject);
begin

    TDBGrid(Sender).DataSource.DataSet.Edit;
    TDBGrid(Sender).DataSource.DataSet.FieldByName('IsSelect').AsInteger:=
       (1 - TDBGrid(Sender).DataSource.DataSet.FieldByName('IsSelect').AsInteger);
    TDBGrid(Sender).DataSource.DataSet.Post;

    TDBGrid(Sender).Refresh;
end;
//----------------------------------------------------------------------------------------
procedure TfSetUpEntities.cdsEntitiesIsSelectGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
    IF Sender.Value = 0 Then Text:='нет'
    Else  Text:='да';
end;
//----------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------
//        E N T I T Y     L I S T S
//----------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------
Function TEntityLists.AddEntity( Value : Variant; AName :  String =''; ImageIndex : Integer =0 ) : TEntityItem;
Begin
   Result:= TEntityItem( TLists(self).AddItem(Value,AName,nil) );
   Result.ImageIndex:=ImageIndex;
End;
//----------------------------------------------------------------------------------------
Function TEntityLists.GetEntityItem( Index : Integer) : TEntityItem;
Begin
     Result:=TEntityItem(inherited Items[Index]);
End;
//----------------------------------------------------------------------------------------
procedure TfSetUpEntities.PopupMenu1Popup(Sender: TObject);
begin

 nUp.Enabled:=cdsEntities.RecNo > 1;
 nDown.Enabled:=cdsEntities.RecNo < cdsEntities.RecordCount;

end;
//----------------------------------------------------------------------------------------
procedure TfSetUpEntities.nUpClick(Sender: TObject);
Var OrderNo , RecNo : Integer;
begin

   cdsEntities.DisableControls;

   OrderNO:=cdsEntities.FieldByName('OrderNo').AsInteger + TControl(Sender).Tag;
   RecNo:=cdsEntities.FieldByName('EntityID').AsInteger;

   cdsEntities.Locate('OrderNo' , OrderNo , []);
   cdsEntities.Edit;
   cdsEntities.FieldByName('OrderNo').AsInteger:=OrderNo -  TControl(Sender).Tag;
   cdsEntities.Post;


   cdsEntities.Locate('EntityID' , RecNo , []);

   cdsEntities.Edit;
   cdsEntities.FieldByName('OrderNo').AsInteger:=OrderNo;
   cdsEntities.Post;

   cdsEntities.EnableControls;

end;
//----------------------------------------------------------------------------------------
procedure TfSetUpEntities.BitBtn2Click(Sender: TObject);
Var i : Integer;
begin
   cdsEntities.First;

   For i:=1 To cdsEntities.RecordCount Do
   Begin
     qCommand.ParamByName('IsSelect').AsInteger:=cdsEntities.FieldByName('IsSelect').AsInteger;
     qCommand.ParamByName('OrderNo').AsInteger:=cdsEntities.FieldByName('OrderNo').AsInteger;
     qCommand.ParamByName('EntityID').AsInteger:=cdsEntities.FieldByName('EntityID').AsInteger;

     qCommand.ExecSql;

     cdsEntities.Next;
   End;


end;
//----------------------------------------------------------------------------------------
Initialization
   SetDataBaseName('TreeDB');
end.
