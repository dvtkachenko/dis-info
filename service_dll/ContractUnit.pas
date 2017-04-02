unit ContractUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, DBGrids, Db, DBTables;

type
  TChooseContractForm = class(TForm)
    ContractPanel: TPanel;
    ChooseBitBtn: TBitBtn;
    CacelBitBtn: TBitBtn;
    ContractDBGrid: TDBGrid;
    ChooseContractDataSource: TDataSource;
    ChooseContractQuery: TQuery;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ChooseContractForm: TChooseContractForm;

implementation

{$R *.DFM}

end.
