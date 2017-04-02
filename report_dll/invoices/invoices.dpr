library invoices;
{ ≈сли сервисы предоставл€емые данной библиотекой будут
  использоватьс€ из других DLL данного приложени€, то они
  должны быть реализованы таким образом , чтобы было
  возможным осуществл€ть их программную инициализацию
  и программный вызов из других DLL , с передачей
  соответствующих параметров
}
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
  invoices_relUnit in 'invoices_relUnit.pas' {InvRelExportForm},
  shared_type in '..\..\shared_type.pas',
  excel_type in '..\..\excel_type.pas',
  invoicesUnit in 'invoicesUnit.pas' {InvExportForm};

Var parentOwner : TApplication;
    parentConfig : p_config;
    InitFlag : boolean;
    // данна€ переменна€ предназначена дл€
    // организации вызова из различных DLL
    // данного приложени€
    InvExportForm: TInvExportForm;

procedure servInitDLL(Owner : TApplication; DB : TDatabase); external 'service.dll' name 'InitDLL';
procedure servUnInitDLL; external 'service.dll' name 'UnInitDLL';

//--------------------------------------
procedure GetMenuName(menu_name : PChar); export;
begin
  StrPCopy(menu_name,'Ёкспорт счетов-фактур');
end;

//--------------------------------------
// процедура предназначена дл€ организации вызовов из различных DLL
// данного приложени€
// данна€ процедура инициализирует необходимую форму и
// передает указатель на эту форму в вызвавшую клиентскую DLL
//--------------------------------------
procedure InitServiceExternalCall(Var _InvExportForm : TInvExportForm); export;
begin
  //
  if InitFlag then begin
    // экспорт счетов - фактур (InvExportForm)
    if (InvExportForm = nil) then begin
      InvExportForm := TInvExportForm.Create(parentOwner);
      InvExportForm.InterprocessCall := true;
      InvExportForm.PathToProgram := parentConfig.PathToProgram;
      InvExportForm.parentConfig := parentConfig;
      _InvExportForm := InvExportForm;
      //
{      if (i1page and param) = 0 then
        InvExportForm.InvPageControl.Pages[0].TabVisible := false;
      if (i2page and param) = 0 then
        InvExportForm.InvPageControl.Pages[1].TabVisible := false;
      if (i3page and param) = 0 then
        InvExportForm.InvPageControl.Pages[2].TabVisible := false;
      if (i4page and param) = 0 then
        InvExportForm.InvPageControl.Pages[3].TabVisible := false;
      //
      InvExportForm.ShowModal;
}
    end;
  end
  else
    raise Exception.Create('Ќе инициализирован модуль');
end;

//--------------------------------------
// данна€ процедура выполн€ет деинициализацию формы
//--------------------------------------
procedure DeInitServiceExternalCall; export;
begin
  //
  if (InvExportForm <> nil) then begin
    InvExportForm.Free;
    InvExportForm := nil;
  end;
  //
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
  InvExportForm: TInvExportForm;
  InvRelExportForm: TInvRelExportForm;
begin
  InvExportForm := nil;
  InvRelExportForm := nil;

  if InitFlag then begin
    try
      // экспорт счетов - фактур (InvExportForm)
      if (i1form and param) <> 0 then begin
        InvExportForm := TInvExportForm.Create(parentOwner);
        InvExportForm.InterprocessCall := false;
        InvExportForm.PathToProgram := parentConfig.PathToProgram;
        InvExportForm.parentConfig := parentConfig;
        //
        if (i1page and param) = 0 then
          InvExportForm.InvPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          InvExportForm.InvPageControl.Pages[1].TabVisible := false;
        if (i3page and param) = 0 then
          InvExportForm.InvPageControl.Pages[2].TabVisible := false;
        if (i4page and param) = 0 then
          InvExportForm.InvPageControl.Pages[3].TabVisible := false;
        if (i5page and param) = 0 then
          InvExportForm.InvPageControl.Pages[4].TabVisible := false;
        //
        InvExportForm.ShowModal;
      end;

      // экспорт св€занных счетов - фактур (InvRelExportForm)
      if (i2form and param) <> 0 then begin
        InvRelExportForm := TInvRelExportForm.Create(parentOwner);
        InvRelExportForm.PathToProgram := parentConfig.PathToProgram;
        //
        if (i1page and param) = 0 then
          InvRelExportForm.InvPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          InvRelExportForm.InvPageControl.Pages[1].TabVisible := false;
        if (i3page and param) = 0 then
          InvRelExportForm.InvPageControl.Pages[2].TabVisible := false;
        //
        InvRelExportForm.ShowModal;
      end;

    finally
      if InvExportForm <> nil then InvExportForm.Free;
      if InvRelExportForm <> nil then InvRelExportForm.Free;
    end;
  end
  else
    raise Exception.Create('Ќе инициализирован модуль');
end;

exports
  GetMenuName name 'GetMenuName',
  InitDLL name 'InitDLL',
  UnInitDLL name 'UnInitDLL',
  ShowForm name 'ShowForm',
  InitServiceExternalCall name 'InitServiceExternalCall',
  DeInitServiceExternalCall name 'DeInitServiceExternalCall';

// вызываетс€ при загрузке и выгрузке модул€ dll
begin
  InitFlag := false;
  parentOwner := nil;
  parentConfig := nil;
  InvExportForm := nil;
end.
