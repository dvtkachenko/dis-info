unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DBTables, ComCtrls, Buttons, Db, DBCtrls, Grids, DBGrids;

type

  TForm1 = class(TForm)
    Button1: TButton;
    BitBtn1: TBitBtn;
    Query1: TQuery;
    DataSource1: TDataSource;
    DBComboBox1: TDBComboBox;
    DBGrid1: TDBGrid;
    DBLookupComboBox1: TDBLookupComboBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    procedure Button1Click(Sender: TObject);
    procedure DBComboBox1Change(Sender: TObject);
    procedure DBLookupComboBox1Click(Sender: TObject);
  private
    { Private declarations }
  public
    pos : integer;
    rel_name : string;
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Button1Click(Sender: TObject);
Var
//  Excel : TExcel;
  size1, i : integer;

begin
  {}
//   Excel := TExcel.Create;
  i := 3;
  size1 := sizeof(integer) + i + 3;
//   Excel.AddWorkBook('d:\program\dis_info\template\nds_report.xlt');
//   Excel.Visible := true;
//   Excel.CopyWorkSheet('source','dfg');
//   Excel.Free;
end;

procedure TForm1.DBComboBox1Change(Sender: TObject);
begin
  pos := Query1.fieldbyname('rel_type_id').asinteger;;
  rel_name := Query1.paramByName('rel_name').asstring;

end;

procedure TForm1.DBLookupComboBox1Click(Sender: TObject);
begin
  pos := Query1.fieldbyname('rel_type_id').asinteger;;
  rel_name := Query1.paramByName('rel_name').asstring;
end;

end.
