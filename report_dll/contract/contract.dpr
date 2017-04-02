library contract;

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
  Forms,
  DBTables,
  serviceDataUnit in '..\..\service_dll\serviceDataUnit.pas' {serviceDataModule: TDataModule},
  shared_type in '..\..\shared_type.pas',
  excel_type in '..\..\excel_type.pas',
  contractUnit in 'contractUnit.pas' {ContractForm},
  contract_statisticUnit in 'contract_statisticUnit.pas' {contract_statisticForm},
  Excel_TLB in '..\..\..\..\..\Program Files\borland\Delphi4\Imports\Excel_TLB.pas',
  Word_TLB in '..\..\..\..\..\Program Files\Borland\Delphi4\Imports\Word_TLB.pas';

Var parentOwner : TApplication;
    parentConfig : p_config;
    InitFlag : boolean;
//    serviceDataModule: TserviceDataModule;

procedure servInitDLL(Owner : TApplication; DB : TDatabase); external 'service.dll' name 'InitDLL';
procedure servUnInitDLL; external 'service.dll' name 'UnInitDLL';

//--------------------------------------
procedure GetMenuName(menu_name : PChar); export;
begin
  StrPCopy(menu_name,'Информация о договорах');
end;

//----------------------------------------------------------------------
procedure InitDLL(Conf : p_config); export;
begin
  if not InitFlag then begin
    parentOwner := Conf.Owner;
    Application.Handle := Conf.Owner.Handle;
    parentConfig := Conf;
//    serviceDataModule := TserviceDataModule.Create(Application);
//    serviceDataModule.serviceDatabaseCyrr.DatabaseName := Config.DBcyrr.DatabaseName;
//    serviceDataModule.serviceDatabaseCyrr.Handle := Config.DBcyrr.Handle;
//    serviceDataModule.serviceDatabaseCyrr.HandleShared := true;
    servInitDLL(parentOwner,Conf.DBcyrr);
    InitFlag := true;
  end;
end;

//----------------------------------------------------------------------
procedure UnInitDLL; export;
begin
  if InitFlag then begin
    servUnInitDLL;
//    serviceDataModule.Free;
    InitFlag := false;
  end;
end;

//--------------------------------------
procedure ShowForm(param : integer); export;
var
  ContractForm: TContractForm;
  contract_statisticForm : Tcontract_statisticForm;
begin
  ContractForm := nil;
  contract_statisticForm := nil;
  if InitFlag then begin
    try
      // форма для вытяжки отчетов о договорах
      if (i1form and param) <> 0 then begin
        ContractForm := TContractForm.Create(parentOwner);
        ContractForm.PathToProgram := parentConfig.PathToProgram;
        if (i1page and param) = 0 then
          ContractForm.contractPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          ContractForm.contractPageControl.Pages[1].TabVisible := false;
        //
        ContractForm.ShowModal;
      end;

      // форма для вытяжки статистики работы по разным группам договоров
      if (i2form and param) <> 0 then begin
        contract_statisticForm := Tcontract_statisticForm.Create(parentOwner);
        contract_statisticForm.PathToProgram := parentConfig.PathToProgram;
        if (i1page and param) = 0 then
          contract_statisticForm.contract_grpPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          contract_statisticForm.contract_grpPageControl.Pages[1].TabVisible := false;
        if (i3page and param) = 0 then
          contract_statisticForm.contract_grpPageControl.Pages[2].TabVisible := false;
        //
        contract_statisticForm.ShowModal;
      end;


    finally
      if ContractForm <> nil then ContractForm.Free;
      if contract_statisticForm <> nil then contract_statisticForm.Free;
    end;
  end
  else
    raise Exception.Create('Не инициализирован модуль');
end;

exports
  GetMenuName name 'GetMenuName',
  InitDLL name 'InitDLL',
  UnInitDLL name 'UnInitDLL',
  ShowForm name 'ShowForm';

// вызывается при загрузке и выгрузке модуля dll
begin
  InitFlag := false;
  parentOwner := nil;
  parentConfig := nil;
end.
