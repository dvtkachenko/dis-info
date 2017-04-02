program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  Excel_TLB in 'C:\Program Files\Borland\Delphi4\Lib\Excel_TLB.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
