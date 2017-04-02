library service;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  SysUtils,
  Classes,
  Controls,
  DBTables,
  Forms,
  BDE,
  findProductionUnit in 'findProductionUnit.pas' {findProductionForm},
  DepatmentUnit in 'DepatmentUnit.pas' {DepatmentForm},
  serviceDataUnit in 'serviceDataUnit.pas' {serviceDataModule: TDataModule},
  ContractUnit in 'ContractUnit.pas' {ChooseContractForm},
  FindEnterpriseUnit in 'FindEnterpriseUnit.pas' {FindEnterpriseForm},
  shared_type in '..\shared_type.pas',
  Excel_TLB in '..\..\..\..\Program Files\borland\Delphi4\Imports\Excel_TLB.pas';

Var parentOwner : TApplication;
    parentDB : TDatabase;
    InitFlag : boolean;
//    serviceDataModule: TserviceDataModule;

//----------------------------------------------------------------------
procedure InitDLL(Owner : TApplication; DB : TDatabase); export;
begin
  if not InitFlag then begin
    parentOwner := Owner;
    Application.Handle := Owner.Handle;
    parentDB := DB;
//    serviceDataModule := TserviceDataModule.Create(Application);
//    serviceDataModule.serviceDatabaseCyrr.DatabaseName := parentDB.DatabaseName;
//    serviceDataModule.serviceDatabaseCyrr.Handle := parentDB.Handle;
//    serviceDataModule.serviceDatabaseCyrr.HandleShared := true;
    InitFlag := true;
  end;
end;

//----------------------------------------------------------------------
procedure UnInitDLL; export;
begin
  if InitFlag then begin
//    serviceDataModule.Free;
    InitFlag := false;
  end;
end;

//----------------------------------------------------------------------
function GetContract(id:integer;Var contract_id:integer;Var contract:PChar) : integer; export;
var
  ChooseContractForm : TChooseContractForm;
  i : integer;
  s : string;
begin
  if InitFlag then begin
    ChooseContractForm := nil;
    try
      ChooseContractForm := TChooseContractForm.Create(parentOwner);
      ChooseContractForm.ChooseContractQuery.ParamByName('ent_id').asfloat := id;
      ChooseContractForm.ChooseContractQuery.Open;
      i := ChooseContractForm.ShowModal;
      if i = mrOk then begin
        s  := ChooseContractForm.ChooseContractQuery.fieldbyname('contract_no').asstring;
        StrPCopy(contract,s);
        contract_id := ChooseContractForm.ChooseContractQuery.fieldbyname('contract_id').asinteger;
      end
      else begin
        contract[0] := #0;
      end;
      ChooseContractForm.ChooseContractQuery.Close;
    finally
      ChooseContractForm.Free;
    end;
  end
  else
    raise Exception.Create('Не инициализирован модуль');
  GetContract := i;
end;

//----------------------------------------------------------------------
function GetDepatment(Var id:integer;name:PChar) : integer; export;
var
  DepatmentForm : TDepatmentForm;
  i : integer;
  s : string;
begin
  if InitFlag then begin
    DepatmentForm := nil;
    try
      DepatmentForm := TDepatmentForm.Create(parentOwner);
      DepatmentForm.DepatmentTable.Open;
      DepatmentForm.SubdivisionTable.Open;
      i := DepatmentForm.ShowModal;
      if i = mrOk then begin
        id := DepatmentForm.SubdivisionTable.fieldbyname('subdivision_id').asinteger;
        s  := DepatmentForm.SubdivisionTable.fieldbyname('subdivision_name').asstring;
        StrPCopy(name,s);
      end
      else begin
        id := 0;
        name[0] := #0;
      end;
      DepatmentForm.DepatmentTable.Close;
      DepatmentForm.SubdivisionTable.Close;
    finally
      DepatmentForm.Free;
    end;
  end
  else
    raise Exception.Create('Не инициализирован модуль');
  GetDepatment := i;
end;

//----------------------------------------------------------------------
function GetEnterprise(Var id:integer;name:PChar): integer; export;
Var
  FindEnterpriseForm : TFindEnterpriseForm;
  i : integer;
  s : string;
begin
  if InitFlag then begin
    FindEnterpriseForm := nil;
    try
      FindEnterpriseForm := TFindEnterpriseForm.Create(parentOwner);
      FindEnterpriseForm.EnterpriseEdit.Text := '';
      i := FindEnterpriseForm.ShowModal;
      id := 0;
      if i = mrOk then begin
        id := FindEnterpriseForm.FindEnterpriseQuery.fieldbyname('object_id').asinteger;
        s := FindEnterpriseForm.FindEnterpriseQuery.fieldbyname('object_name').asstring;
        if (id = 0) and (s = '') then begin
          i := mrCancel;
          name[0] := #0;
        end;
        StrPCopy(name,s);
        FindEnterpriseForm.FindEnterpriseQuery.Close;
      end
      else begin
        id := 0;
        name[0] := #0;
      end;
    finally
      FindEnterpriseForm.Free;
    end;
  end
  else
    raise Exception.Create('Не инициализирован модуль');
  GetEnterprise := i;
end;

//----------------------------------------------------------------------
function GetProduction(const mode:integer ; Var id:integer; name:PChar): integer; export;
Var
  findProductionForm : TfindProductionForm;
  i : integer;
  s : string;
begin
  if InitFlag then begin
    findProductionForm := nil;
    try
      findProductionForm := TfindProductionForm.Create(parentOwner);
      findProductionForm.productionEdit.Text := '';

      // формируем запрос в зависимости от режима
      // поиска или по продукции или по ее типу
      // или по доп.расходам

      if mode = iprod_mode then begin
        with findProductionForm.findProductionQuery do begin
          Close;
          SQL.Clear;
          SQL.Add('select supply_id id, trade_mark name');
          SQL.Add('from supply');
          SQL.Add('where upper(trade_mark) like upper(:param)');
          SQL.Add('order by trade_mark');
          Prepare;
        end;
        findProductionForm.Caption := 'Поиск по наименованию продукции';
        findProductionForm.productionDBGrid.Columns[1].Title.Caption :=
            'Наименование продукции';
      end;

      if mode = iprod_type_mode then begin
        with findProductionForm.findProductionQuery do begin
          Close;
          SQL.Clear;
          SQL.Add('select prod_id id, prod_name name');
          SQL.Add('from products');
          SQL.Add('where upper(prod_name) like upper(:param)');
          SQL.Add('order by prod_name');
          Prepare;
        end;
        findProductionForm.Caption := 'Поиск по типу продукции';
        findProductionForm.productionDBGrid.Columns[1].Title.Caption :=
            'Тип продукции';
      end;

      if mode = iextra_item_mode then begin
        with findProductionForm.findProductionQuery do begin
          Close;
          SQL.Clear;
          SQL.Add('select extra_id id, extra_item_name name');
          SQL.Add('from extra_items_guide');
          SQL.Add('where upper(extra_item_name) like upper(:param)');
          SQL.Add('order by extra_item_name');
          Prepare;
        end;
        findProductionForm.Caption := 'Поиск по наименованию доп. расходов';
        findProductionForm.productionDBGrid.Columns[1].Title.Caption :=
            'Наименование доп.расходов';
      end;
      // конец формирования запроса

      i := findProductionForm.ShowModal;
      id := 0;
      if i = mrOk then begin
        id := findProductionForm.findProductionQuery.fieldbyname('id').asinteger;
        s := findProductionForm.findProductionQuery.fieldbyname('name').asstring;
        if (id = 0) and (s = '') then begin
          i := mrCancel;
          name[0] := #0;
        end;
        StrPCopy(name,s);
        findProductionForm.findProductionQuery.Close;
      end
      else begin
        id := 0;
        name[0] := #0;
      end;
    finally
      findProductionForm.Free;
    end;
  end
  else
    raise Exception.Create('Не инициализирован модуль');
  GetProduction := i;
end;

exports
  InitDLL name 'InitDLL',
  GetContract name 'GetContract',
  GetDepatment name 'GetDepatment',
  GetEnterprise name 'GetEnterprise',
  GetProduction name 'GetProduction',
  UnInitDLL name 'UnInitDLL';

{procedure InitDBTables;
begin
  if SaveInitProc <> nil then TProcedure(SaveInitProc);
  NeedToUninitialize := Succeeded(CoInitialize(nil));
end;

initialization
  if not IsLibrary then
  begin
    SaveInitProc := InitProc;
    InitProc := @InitDBTables;
  end;
//finalization
//  DepatmentForm.Free;
//  FindEnterpriseForm.Free;}
begin
  parentOwner := nil;
  InitFlag := false;
//  serviceDataModule := serviceDataModule;
//  serviceDataModule.serviceDatabaseCyrr.Handle := serviceDataModule.serviceDatabaseCyrr.Handle;
//  serviceDataModule.serviceDatabaseCyrr.HandleShared := true;
//if reason = DLL_PROCESS_ATTACH
//      DMod := TDMod.Create(Nil);
end.

