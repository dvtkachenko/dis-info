program dis_info;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {DISMainForm},
  MainDataUnit in 'MainDataUnit.pas' {MainDataModule: TDataModule},
  ZastUnit in 'ZastUnit.pas' {ZastForm},
  main_type in 'main_type.pas',
  shared_type in 'shared_type.pas',
  prikol_form in 'prikol_form.pas' {prikolForm},
  Excel_TLB in '..\..\..\Program Files\Borland\Delphi4\Imports\Excel_TLB.pas';

Var
  ZastForm : TZastForm;

{$R *.RES}

begin
 try
  Application.Initialize;
  Application.Title := 'ООО "ДИС"';
  ZastForm := TZastForm.Create(Application);
  ZastForm.Show;
  ZastForm.Refresh;
  Application.CreateForm(TMainDataModule, MainDataModule);
  ZastForm.Refresh;
  Application.CreateForm(TDISMainForm, DISMainForm);
  Application.CreateForm(TprikolForm, prikolForm);
  // активизируем соединение с базой только после создания формы
  // т.к. до этого еще не установлено св-во TSession.NetFileDir
//  with MainDataModule do begin
//    DatabaseDIS.Connected := true;
//    DatabaseDIScyrr.Connected := true;
//  end;
  ZastForm.Refresh;
  ZastForm.Close;

  DISMainForm.prikolTimer.Enabled := false;
  DISMainForm.prikolTimer.tag := 1;
  if (DISMainForm.Config.conf.username = 'Yvoleynik') then begin
//    prikolForm.canJoke := true;
//    DISMainForm.prikolTimer.Enabled := true;
  end;

  Application.Run;
  finally
    if ZastForm <> nil then ZastForm.Free;
//    with MainDataModule do begin
//      DatabaseDIS.Connected := false;
//      DatabaseDIScyrr.Connected := false;
//    end;
  end;
end.
