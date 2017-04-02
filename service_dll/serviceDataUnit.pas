unit serviceDataUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables;

type
  TserviceDataModule = class(TDataModule)
    serviceDatabaseCyrr: TDatabase;
    serviceDatabase: TDatabase;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}


end.

