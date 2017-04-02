//----------------------------------------------------------------------------------------
//  "DB Дерево своими руками"
//
// Специально для Королевства Delphi http://www.delphikingdom.com
//
// Елена Филиппова , 27 марта 2001 г.
//
//----------------------------------------------------------------------------------------
unit EditUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, StdCtrls, Buttons, DBCtrls, Provider, DBClient;

type
  TFormEdit = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    edName: TEdit;
    listParent: TDBLookupComboBox;
    BitBtn2: TBitBtn;
    BitBtn1: TBitBtn;
    DataSource1: TDataSource;
    qList: TQuery;
    qCommand: TQuery;
    cdsList: TClientDataSet;
    DataSetProvider1: TDataSetProvider;
    qListName: TStringField;
    qListID: TIntegerField;
    cdsListName: TStringField;
    cdsListID: TIntegerField;
    procedure BitBtn2Click(Sender: TObject);
    procedure cdsListAfterOpen(DataSet: TDataSet);
  private
    isInsert  : Boolean;
    CompanyID : Integer;

    Procedure PrepareUpdate( ID : Integer);
  public
    { Public declarations }
  end;

var
  FormEdit: TFormEdit;

Function ModifyCompany( ID : Integer = 0; ParentID : Integer = 0 ) : Boolean;

implementation

{$R *.DFM}
Const
   SqlInfo   = 'SELECT Name , ParentID FROM Company WHERE ID=%d';
   SqlModify : array [0..1] of String
                 = ('UPDATE  Company SET Name = ''%s'' WHERE ID = %d ',
                    'INSERT INTO Company ( Name , ParentID ) VALUES(''%s'' , %d)');
//----------------------------------------------------------------------------------------
Function ModifyCompany( ID : Integer = 0; ParentID : Integer = 0 ) : Boolean;
Begin

    IF  FormEdit = nil Then Application.CreateForm(TFormEdit,FormEdit);

    With FormEdit Do
    Begin
       IsInsert:=( ID = 0 );
       CompanyID:=ID;

       cdsList.Close;
       cdsList.Open;

       listParent.Enabled:=IsInsert;

       IF NOT isInsert Then PrepareUpdate(ID)
       Else Begin
               EdName.Text:='';
               listParent.KeyValue:=ParentID;
            End;

       Result:= (ShowModal = mrOk);
    End;

End;
//----------------------------------------------------------------------------------------
Procedure TFormEdit.PrepareUpdate( ID : Integer);
Begin
       qCommand.Close;
       qCommand.Sql.Clear;
       qCommand.Sql.Add(Format(SqlInfo,[ID]));
       qCommand.Open;

       EdName.Text:=qCommand.FieldByName('Name').AsString;
       listParent.KeyValue:=qCommand.FieldByName('ParentID').AsInteger;

End;
//----------------------------------------------------------------------------------------
procedure TFormEdit.BitBtn2Click(Sender: TObject);
Var ID : Integer;
begin
   IF edName.Text = '' Then MessageDlg( 'Не задано имя' , mtError , [mbOk] , 0 )
   Else Try
             IF IsInsert Then ID:=cdsList.FieldByName('ID').AsInteger
             Else ID:=CompanyID;

             qCommand.Close;
             qCommand.Sql.Clear;
             qCommand.Sql.Add(Format(SqlModify[ORD(isInsert)],[edName.Text , ID ]));
             qCommand.ExecSql;

             ModalResult:=mrOk;
         Except
              On E : Exception Do Exception.Create('Ошибка записи - ' + E.Message);
         End;

end;
//----------------------------------------------------------------------------------------
procedure TFormEdit.cdsListAfterOpen(DataSet: TDataSet);
begin
      DataSet.InsertRecord(['< Пусто >',0]);
end;
//----------------------------------------------------------------------------------------

end.
