unit MainDataUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables;

type
  TMainDataModule = class(TDataModule)
    DatabaseISD2000: TDatabase;
    DatabaseDIS: TDatabase;
    DatabaseDIScyrr: TDatabase;
    Database_dis_ibdb_cyrr: TDatabase;
    procedure MainDataModuleCreate(Sender: TObject);
    procedure MainDataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainDataModule: TMainDataModule;

implementation

uses MainUnit;

{$R *.DFM}


procedure TMainDataModule.MainDataModuleCreate(Sender: TObject);
begin
//  DatabaseNew.Connected := true;
  DatabaseDIS.Connected := true;
//  DatabaseDIScyrr.Connected := true;
//  DatabaseISD2000.Connected := true;
end;

procedure TMainDataModule.MainDataModuleDestroy(Sender: TObject);
begin
//  DatabaseNew.Connected := false;
  DatabaseDIS.Connected := false;
//  DatabaseDIScyrr.Connected := false;
//  DatabaseISD2000.Connected := false;
end;

end.
