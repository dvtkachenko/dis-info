unit prikol_form;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TprikolForm = class(TForm)
    prikolLabel: TLabel;
    prikolBitBtn: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure prikolBitBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    canJoke : boolean;
    keepShow : boolean;
  end;

var
  prikolForm: TprikolForm;

implementation

{$R *.DFM}

procedure TprikolForm.FormCreate(Sender: TObject);
begin
  canJoke := false;
  keepShow := false;
end;

procedure TprikolForm.prikolBitBtnClick(Sender: TObject);
begin
  Close;
  keepShow := false;
end;

end.
