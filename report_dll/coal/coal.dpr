library coal;

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
  coalUnit in 'coalUnit.pas' {coalReportForm},
  shared_type in '..\..\shared_type.pas',
  excel_type in '..\..\excel_type.pas',
  Excel_TLB in '..\..\..\..\..\Program Files\borland\Delphi4\Imports\Excel_TLB.pas';

Var parentOwner : TApplication;
    parentConfig : p_config;
    InitFlag : boolean;
//    serviceDataModule: TserviceDataModule;

procedure servInitDLL(Owner : TApplication; DB : TDatabase); external 'service.dll' name 'InitDLL';
procedure servUnInitDLL; external 'service.dll' name 'UnInitDLL';

//--------------------------------------
procedure GetMenuName(menu_name : PChar); export;
begin
  StrPCopy(menu_name,'разбивка углей');
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
  CoalReportForm: TCoalReportForm;
begin
  CoalReportForm := nil;
  if InitFlag then begin
    try
      coalReportForm := TcoalReportForm.Create(parentOwner);
      coalReportForm.PathToProgram := parentConfig.PathToProgram;

      if (i1page and param) = 0 then
        CoalReportForm.mainPageControl.Pages[0].TabVisible := false;

      coalReportForm.ShowModal;
    finally
      coalReportForm.Free;
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
