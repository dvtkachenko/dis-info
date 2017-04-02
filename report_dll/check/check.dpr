library check;

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
  checkUnit in 'checkUnit.pas' {checkDataForm},
  check_dis_isdUnit in 'check_dis_isdUnit.pas' {comp_dis_isdForm},
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
  StrPCopy(menu_name,'Проверка правильности данных');
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
  checkDataForm: TcheckDataForm;
  comp_dis_isdForm: Tcomp_dis_isdForm;
begin
  checkDataForm := nil;
  comp_dis_isdForm := nil;

  if InitFlag then begin
    try
      // проверка корректности данных в БД ДИС98 (checkDataForm)
      if (i1form and param) <> 0 then begin
        checkDataForm := TcheckDataForm.Create(parentOwner);
        checkDataForm.PathToProgram := parentConfig.PathToProgram;
        checkDataForm.parentConfig := parentConfig;
        //
        if (i1page and param) = 0 then
          checkDataForm.checkDataPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          checkDataForm.checkDataPageControl.Pages[1].TabVisible := false;
//        if (i3page and param) = 0 then
//          checkDataForm.InvPageControl.Pages[2].TabVisible := false;
//        if (i4page and param) = 0 then
//          checkDataForm.InvPageControl.Pages[3].TabVisible := false;
        //
        checkDataForm.ShowModal;
      end;

      // сверка взаиморасчетов между ДИС и ИСД
      // по базам данных dis98 и isd2000 соответственно
      if (i2form and param) <> 0 then begin
        parentConfig.ora_isdDB.Connected := true;

        comp_dis_isdForm := Tcomp_dis_isdForm.Create(parentOwner);
        comp_dis_isdForm.PathToProgram := parentConfig.PathToProgram;
        //
        if (i1page and param) = 0 then
          comp_dis_isdForm.comp_dis_isdPageControl.Pages[0].TabVisible := false;
        //
        comp_dis_isdForm.ShowModal;
      end;

    finally
      if checkDataForm <> nil then checkDataForm.Free;
      if comp_dis_isdForm <> nil then begin
        comp_dis_isdForm.Free;
        parentConfig.ora_isdDB.Connected := false;
      end;  
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
