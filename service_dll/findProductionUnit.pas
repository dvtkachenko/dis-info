unit findProductionUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Grids, DBGrids, Buttons, Db, DBTables;

type
  TfindProductionForm = class(TForm)
    EnterprisePanel: TPanel;
    FindLabel: TLabel;
    productionEdit: TEdit;
    productionDBGrid: TDBGrid;
    ChooseBitBtn: TBitBtn;
    CancelBitBtn: TBitBtn;
    FindBitBtn: TBitBtn;
    findProductionDataSource: TDataSource;
    findProductionQuery: TQuery;
    procedure FindBitBtnClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure productionEditEnter(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}

procedure TfindProductionForm.FindBitBtnClick(Sender: TObject);
begin
  findProductionQuery.Close;
  findProductionQuery.ParamByName('param').asstring :=
                           '%'+productionEdit.Text+'%';
  findProductionQuery.Open;
end;

procedure TfindProductionForm.FormActivate(Sender: TObject);
begin
  productionEdit.SetFocus;
end;

procedure TfindProductionForm.productionEditEnter(Sender: TObject);
begin
  findProductionQuery.Close;
end;

end.
