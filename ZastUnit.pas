unit ZastUnit;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls;

type
  PZastForm = ^TZastForm;     
  TZastForm = class(TForm)
    ZastPanel: TPanel;
    ZastImage: TImage;
    procedure FormClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}

procedure TZastForm.FormClick(Sender: TObject);
begin
  Close;
end;

end.

