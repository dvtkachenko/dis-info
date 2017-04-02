unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Menus, DBTables, service_type, ActnList,
  ComCtrls, xDBTree, Db, Buttons, ToolWin, ImgList;

type
  TCoolForm = class(TForm)
    mainStatusBar: TStatusBar;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbShowForm: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    ToolButton1: TToolButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Config : TConfig;
  end;

var
  CoolForm: TCoolForm;

implementation

{$R *.DFM}

end.
