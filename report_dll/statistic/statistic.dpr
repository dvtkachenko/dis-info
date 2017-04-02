library statistic;

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
  statisticSaldoCloseUnit in 'statisticSaldoCloseUnit.pas' {statisticSaldoCloseForm},
  shared_type in '..\..\shared_type.pas',
  excel_type in '..\..\excel_type.pas',
  StatisticReportUnit in 'StatisticReportUnit.pas' {StatisticReportForm},
  invoicesUnit in '..\invoices\invoicesUnit.pas' {InvExportForm};

Var parentOwner : TApplication;
    parentConfig : p_config;
    InitFlag : boolean;
//    serviceDataModule: TserviceDataModule;

procedure servInitDLL(Owner : TApplication; DB : TDatabase); external 'service.dll' name 'InitDLL';
procedure servUnInitDLL; external 'service.dll' name 'UnInitDLL';

//--------------------------------------
procedure GetMenuName(menu_name : PChar); export;
begin
  StrPCopy(menu_name,'Экспорт счетов-фактур');
end;

//----------------------------------------------------------------------
procedure InitDLL(Conf : p_config); export;
begin
  if not InitFlag then begin
    parentOwner := Conf.Owner;
    Application.Handle := Conf.Owner.Handle;
    parentConfig := Conf;
//    serviceDataModule := TserviceDataModule.Create(Application);
//    serviceDataModule.serviceDatabase.DatabaseName := Config.DB.DatabaseName;
//    serviceDataModule.serviceDatabase.Handle := Config.DB.Handle;
//    serviceDataModule.serviceDatabase.HandleShared := true;
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
  StatisticForm: TStatisticReportForm;
  statisticSaldoCloseForm : TstatisticSaldoCloseForm;
begin
  StatisticForm := nil;
  statisticSaldoCloseForm := nil;
  if InitFlag then begin
    try
      // формирование различной статистики (StatisticForm)
      if (i1form and param) <> 0 then begin
        StatisticForm := TStatisticReportForm.Create(parentOwner);
        StatisticForm.PathToProgram := parentConfig.PathToProgram;
        StatisticForm.parentConfig := parentConfig;
        //
        if (i1page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[1].TabVisible := false;
        if (i3page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[2].TabVisible := false;
        if (i4page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[3].TabVisible := false;
        if (i5page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[4].TabVisible := false;
        if (i6page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[5].TabVisible := false;
        if (i7page and param) = 0 then
          StatisticForm.StatisticPageControl.Pages[6].TabVisible := false;
        //
        StatisticForm.ShowModal;
      end;

      // экспорт статистики по передприятиям
      // по которым возможно закрытие
      // по встречным договорам  (statisticSaldoCloseForm)
      if (i2form and param) <> 0 then begin
        statisticSaldoCloseForm := TstatisticSaldoCloseForm.Create(parentOwner);
        statisticSaldoCloseForm.PathToProgram := parentConfig.PathToProgram;
        //
        if (i1page and param) = 0 then
          statisticSaldoCloseForm.StatisticPageControl.Pages[0].TabVisible := false;
        //
        statisticSaldoCloseForm.ShowModal;
      end;

    finally
      if StatisticForm <> nil then StatisticForm.Free;
      if statisticSaldoCloseForm <> nil then statisticSaldoCloseForm.Free;
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
