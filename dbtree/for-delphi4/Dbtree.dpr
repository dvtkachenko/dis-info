program DBTree;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  dlgTree in 'Dlgtree.pas' {dlgTreeEdit};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
