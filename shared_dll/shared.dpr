library shared;

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
  shared_res in 'shared_res.pas',
  shared_type in '..\shared_type.pas';

Var busy, InitFlag : boolean;
    BeginDate, EndDate : TDateTime;

//----------------------------------------------------------------------
procedure InitDLL; export;
begin
  if not InitFlag then begin
    InitFlag := true;
  end;
end;

//----------------------------------------------------------------------
procedure UnInitDLL; export;
begin
  if InitFlag then begin
    InitFlag := false;
  end;
end;

//----------------------------------------------------------------------
function ReadDate(Var _BeginDate, _EndDate : TDateTime):boolean; export;
begin
  if not busy then begin
    busy := true;
    _BeginDate := BeginDate;
    _EndDate := EndDate;
    ReadDate := true;
    busy := false;
  end
  else
    ReadDate := false;
end;

//----------------------------------------------------------------------
function WriteDate(_BeginDate, _EndDate : TDateTime):boolean; export;
begin
  if not busy then begin
    busy := true;
    BeginDate := _BeginDate;
    EndDate := _EndDate;
    WriteDate := true;
    busy := false;
  end
  else
    WriteDate := false;
end;

exports
  InitDLL name 'InitDLL',
  UnInitDLL name 'UnInitDLL',
  ReadDate name 'ReadDate',
  WriteDate name 'WriteDate';

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
  busy := false;
  InitFlag := true;
  BeginDate := StrToDate(startDate);
  EndDate := Date;
end.

