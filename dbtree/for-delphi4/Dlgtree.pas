{
  Copyright © 1996-1997 TG Credit Bank, Inc.,
  Yaroslav Steshko    (yari@tgbank.dol.ru)
  Alexander Buzaev    (buzaev@sbank.ru)

  ”правление параметрами веток дерева. ѕозвол€ет:
   - создавать и измен€ть ветки и их названи€;
   - индекс plugin'a (FormID);
   - картинки, отображаемые в древе.
}
unit dlgTree;

interface

uses
  Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, Grids, DBGrids, DB, DBTables, Dialogs, ComCtrls,
  Mask, DBCtrls, ExtDlgs;

type
  TdlgTreeEdit = class(TForm)
    bnOk: TButton;
    bnCancel: TButton;
    dsTree: TDataSource;
    qrTree: TQuery;
    Label3: TLabel;
    DBEdit1: TDBEdit;
    DBRadioGroup1: TDBRadioGroup;
  private
  protected
    OwnerID, ID, PicID: Integer;
  public
  end;

function ExecTreeDlg(AnID, AnOwnerID: Integer): Boolean;

implementation

uses Main;

{$R *.DFM}

function ExecTreeDlg(AnID, AnOwnerID: Integer): Boolean;
begin
  result:=False;
  with TdlgTreeEdit.Create(Application) do
  try
    ID:=AnID;
    OwnerID:=AnOwnerID;
    with qrTree do begin
      SQL.Clear;
      SQL.Add('SELECT * FROM DBTree');
      SQL.Add('WHERE ID = ' + IntToStr(ID));
      Open;
     if ID < 0 then
     begin
       ID:=Main.GetMaxID('DBTree');
       Append;
       FieldByName('ID').AsInteger:=ID;
       if OwnerID > 0 then
         FieldByName('OwnerID').AsInteger := OwnerID;
       FieldByName('PictureID').AsInteger := -1;
     end else
       Edit;
    end;
    if ShowModal = mrOK then begin
      qrTree.Post;
      result:=True;
    end else
      qrTree.Cancel; 
  finally
    Free;
  end;
end;

end.
