program TreeDB;

uses
  Forms,
  TreeUnit in 'TreeUnit.pas' {FormTree},
  EditUnit in 'EditUnit.pas' {FormEdit},
  setupEntities in 'setupEntities.pas' {fSetUpEntities},
  DBManagerUtils in 'DBManagerUtils.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TFormTree, FormTree);
  Application.CreateForm(TfSetUpEntities, fSetUpEntities);
  Application.CreateForm(TFormEdit, FormEdit);

  Application.Run;
end.
