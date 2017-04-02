unit FindEnterpriseUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Grids, DBGrids, Buttons, Db, DBTables;

type
  TFindEnterpriseForm = class(TForm)
    EnterprisePanel: TPanel;
    FindLabel: TLabel;
    EnterpriseEdit: TEdit;
    EnterpriseDBGrid: TDBGrid;
    ChooseBitBtn: TBitBtn;
    CancelBitBtn: TBitBtn;
    FindBitBtn: TBitBtn;
    FindEnterpriseQuery: TQuery;
    FindEnterpriseDataSource: TDataSource;
    procedure FindBitBtnClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure EnterpriseEditEnter(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}

procedure TFindEnterpriseForm.FindBitBtnClick(Sender: TObject);
begin
  with FindEnterpriseQuery do begin
    Close;
    sql.clear;
    sql.add('select distinct enterprise_name object_name, char_okpo, Enterpr_ID object_id from enterpr where ');
    sql.add('upper(enterprise_name collate pxw_cyrl) like upper('+#39+'%' + EnterpriseEdit.Text + '%'+#39+' collate pxw_cyrl)');
    Open;
  end
//  FindEnterpriseQuery.ParamByName('Param1').asstring :=
//                           '%' + EnterpriseEdit.Text + '%' + 'collate pxw_cyrl';
end;

procedure TFindEnterpriseForm.FormActivate(Sender: TObject);
begin
  EnterpriseEdit.SetFocus;
end;

procedure TFindEnterpriseForm.EnterpriseEditEnter(Sender: TObject);
begin
  FindEnterpriseQuery.Close;
end;

end.
