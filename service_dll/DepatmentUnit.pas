unit DepatmentUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, Db, DBTables, StdCtrls, Buttons, ExtCtrls;

type
  TDepatmentForm = class(TForm)
    DepDataSource: TDataSource;
    SubdivDataSource: TDataSource;
    DepatmentTable: TTable;
    SubdivisionTable: TTable;
    OkBitBtn: TBitBtn;
    CancelBitBtn: TBitBtn;
    DepBevel: TBevel;
    DepatmentDBGrid: TDBGrid;
    SubdivisionDBGrid: TDBGrid;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

uses serviceDataUnit;

{$R *.DFM}

end.
