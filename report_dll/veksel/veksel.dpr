library veksel;

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
  vekselUnit in 'vekselUnit.pas' {VekselExportForm},
  shared_type in '..\..\shared_type.pas',
  excel_type in '..\..\excel_type.pas',
  veksel_isdUnit in 'veksel_isdUnit.pas' {Veksel_isdExportForm},
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
  StrPCopy(menu_name,'�������');
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
  VekselExportForm: TVekselExportForm;
  Veksel_isdExportForm: TVeksel_isdExportForm;
begin
  VekselExportForm := nil;
  Veksel_isdExportForm := nil;
  if InitFlag then begin
    try
      // �������� �������� � �� ���98 (VekselExportForm)
      if (i1form and param) <> 0 then begin
        VekselExportForm := TVekselExportForm.Create(parentOwner);
        VekselExportForm.PathToProgram := parentConfig.PathToProgram;
        if (i1page and param) = 0 then
          VekselExportForm.VekselPageControl.Pages[0].TabVisible := false;
        if (i2page and param) = 0 then
          VekselExportForm.VekselPageControl.Pages[1].TabVisible := false;
        if (i3page and param) = 0 then
          VekselExportForm.VekselPageControl.Pages[2].TabVisible := false;
        if (i4page and param) = 0 then
          VekselExportForm.VekselPageControl.Pages[3].TabVisible := false;
        if (i5page and param) = 0 then
          VekselExportForm.VekselPageControl.Pages[4].TabVisible := false;
        //
        VekselExportForm.ShowModal;
      end;

      // �������� �������� � �� ���2000 (Veksel_isdExportForm)
      if (i2form and param) <> 0 then begin
        parentConfig.ora_isdDB.Connected := true;

        Veksel_isdExportForm := TVeksel_isdExportForm.Create(parentOwner);
        Veksel_isdExportForm.PathToProgram := parentConfig.PathToProgram;
        if (i1page and param) = 0 then
          Veksel_isdExportForm.Veksel_isdPageControl.Pages[0].TabVisible := false;
        //
        Veksel_isdExportForm.ShowModal;
      end;

    finally
      if VekselExportForm <> nil then VekselExportForm.Free;
      if Veksel_isdExportForm <> nil then begin
        Veksel_isdExportForm.Free;
        parentConfig.ora_isdDB.Connected := false;
      end;
    end;
  end
  else
    raise Exception.Create('�� ��������������� ������');
end;

exports
  GetMenuName name 'GetMenuName',
  InitDLL name 'InitDLL',
  UnInitDLL name 'UnInitDLL',
  ShowForm name 'ShowForm';

// ���������� ��� �������� � �������� ������ dll
begin
  InitFlag := false;
  parentOwner := nil;
  parentConfig := nil;
end.
