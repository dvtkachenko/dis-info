unit excel_type;

interface
uses SysUtils, Classes, Menus, Windows, Forms, ComObj, OleCtrls, Excel_TLB;

type

  TExcel = class
  private
    xla : _Application;
    xlw : _Workbook;
    LCID : integer;
    procedure fSetVisible(Visible : boolean);
    function fGetVisible : boolean;
    procedure fSetCellFormula(Cell : string; Value : OLEVariant);
    function fGetCellFormula(Cell : string) : OleVariant;
    procedure fSetCell(Cell : string; Value : OLEVariant);
    function fGetCell(Cell : string) : OleVariant;
    procedure fSetRange(CellFrom,CellTo : OleVariant ; Value : Variant);
    function fGetRange(CellFrom,CellTo : OleVariant) : Variant;
  public
//    xla : _Application;
    constructor Create;
    destructor Destroy; override;
    procedure SelectWorkSheet(name : string);
    procedure AddWorkBook(Template : OleVariant);
    procedure SaveAs(filename : string);
    procedure BottomBordersLine(RangeLeft, RangeRight, WorkSheetName : string);
    procedure FillRangeColor(RangeLeft, RangeRight : string; Color : integer);
    procedure RangeFontBold(RangeLeft, RangeRight, WorkSheetName : string);
    procedure CopyWorkSheet(nameSource, nameDest : string);
    property Visible : boolean read fGetVisible write fSetVisible;
    property Cell[Cell : string] : OleVariant read fGetCell write fSetCell;
    property CellFormulaR1C1[Cell : string] : OleVariant read fGetCellFormula write fSetCellFormula;
    property Range[CellFrom,CellTo : OleVariant] : Variant read fGetRange write fSetRange;
  end;


implementation

// ---------------------------------------
// --- реализация методов класса TExcel
// ---------------------------------------
constructor TExcel.Create;
begin
  inherited;
  LCID := GetUserDefaultLCID;
//  xla := CoExcelApplication.Create;
  xla := CoApplication_.Create;
end;
//---------------------------------------
destructor TExcel.Destroy;
begin
  if Assigned(xla) then begin // а если он не создан?
    xla.Quit;
    xla := nil;
  end;
  inherited;
end;
//---------------------------------------
procedure TExcel.AddWorkBook(Template : OleVariant);
begin
  if Assigned(xla) then begin // а если он не создан?
    xlw := xla.Workbooks.Add(Template, LCID);
  end;
end;
//---------------------------------------
procedure TExcel.SelectWorkSheet(name : string);
Var
//  xls : Excel97._Worksheet;
  xls : Excel_TLB._Worksheet;
begin
  if xlw <> nil then begin
//    xls := xlw.Worksheets.Item[name] as Excel97._WorkSheet;
    xls := xlw.Worksheets.Item[name] as Excel_TLB._WorkSheet;
    OleVariant(xls).Select;
    xls := nil;
  end;
end;

//---------------------------------------
procedure TExcel.BottomBordersLine(RangeLeft, RangeRight, WorkSheetName : string);
Var
//  xls : Excel97._Worksheet;
  xls : Excel_TLB._Worksheet;
//  range : Excel97.Range;
  range : Excel_TLB.Range;
begin
  if xlw <> nil then begin
//    xls := xlw.Worksheets.Item[WorkSheetName] as Excel97._WorkSheet;
    xls := xlw.Worksheets.Item[WorkSheetName] as Excel_TLB._WorkSheet;
    xls.Range[OleVariant(RangeLeft), OleVariant(RangeRight)].Select;
//    range := (xla.Selection[lcid] as Excel97.Range);
    range := (xla.Selection[lcid] as Excel_TLB.Range);
    Range.Borders[xlEdgeBottom].LineStyle := xlContinuous;
    Range.Borders[xlEdgeBottom].Weight := xlMedium;
    Range.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
    range := nil;
    xls := nil;
  end;
end;

//---------------------------------------
procedure TExcel.RangeFontBold(RangeLeft, RangeRight, WorkSheetName : string);
Var
//  xls : Excel97._Worksheet;
//  range : Excel97.Range;
  xls : Excel_TLB._Worksheet;
  range : Excel_TLB.Range;
begin
  if xlw <> nil then begin
//    xls := xlw.Worksheets.Item[WorkSheetName] as Excel97._WorkSheet;
    xls := xlw.Worksheets.Item[WorkSheetName] as Excel_TLB._WorkSheet;
    xls.Range[OleVariant(RangeLeft), OleVariant(RangeRight)].Select;
//    range := (xla.Selection[lcid] as Excel97.Range);
    range := (xla.Selection[lcid] as Excel_TLB.Range);
    Range.Font.Bold := True;
    range := nil;
    xls := nil;
  end;
end;
//---------------------------------------
procedure TExcel.FillRangeColor(RangeLeft, RangeRight : string; Color : integer);
Var
//  xls : Excel97._Worksheet;
//  range : Excel97.Range;
  xls : Excel_TLB._Worksheet;
  range : Excel_TLB.Range;
begin
  if xlw <> nil then begin
//    xls := xlw.ActiveSheet as Excel97._WorkSheet;
    xls := xlw.ActiveSheet as Excel_TLB._WorkSheet;
    xls.Range[OleVariant(RangeLeft), OleVariant(RangeRight)].Select;
//    range := (xla.Selection[lcid] as Excel97.Range);
    range := (xla.Selection[lcid] as Excel_TLB.Range);
    Range.Interior.ColorIndex := Color;
    Range.Interior.Pattern := xlSolid;
    range := nil;
    xls := nil;
  end;
end;
//---------------------------------------
procedure TExcel.CopyWorkSheet(nameSource, nameDest : string);
Var
//  xls_source, xls_dest : Excel97._Worksheet;
  xls_source, xls_dest : Excel_TLB._Worksheet;
  xls_OLE : OleVariant;
begin
  if xlw <> nil then begin
//    xls_source := xlw.Worksheets.Item[nameSource] as Excel97._WorkSheet;
    xls_source := xlw.Worksheets.Item[nameSource] as Excel_TLB._WorkSheet;
    xls_OLE := xls_source;
    xls_OLE.Copy(After := xlw.Worksheets.Item[nameSource]);
//    xls_dest := xlw.Worksheets.Item[nameSource + ' (2)'] as Excel97._WorkSheet;
    xls_dest := xlw.Worksheets.Item[nameSource + ' (2)'] as Excel_TLB._WorkSheet;
    xls_dest.name := nameDest;
    xls_source := nil;
    xls_dest := nil;
  end;
end;
//---------------------------------------
procedure TExcel.fSetVisible(Visible : boolean);
begin
  if Assigned(xla) then begin // а если он не создан?
    xla.Visible[lcid] := Visible;
    if Visible then begin
      if xla.WindowState[0] = TOleEnum(xlMinimized) then
        xla.WindowState[0] := TOleEnum(xlNormal);
      xla.ScreenUpdating[0] := true;
    end;
  end;
end;
//---------------------------------------
function TExcel.fGetVisible : boolean;
begin
  result := xla.Visible[lcid];
end;
//---------------------------------------
procedure TExcel.fSetCellFormula(Cell : string; Value : OLEVariant);
begin
  xla.Range[Cell,Cell].FormulaR1C1 := value;
end;
//---------------------------------------
function TExcel.fGetCellFormula(Cell : string) : OLEVariant;
begin
  result := xla.Range[Cell,Cell].FormulaR1C1;
end;
//---------------------------------------
procedure TExcel.fSetCell(Cell : string; Value : OLEVariant);
begin
  xla.Range[Cell,Cell].Value := value;
end;
//---------------------------------------
function TExcel.fGetCell(Cell : string) : OLEVariant;
begin
  result := xla.Range[Cell,Cell].Value;
end;
//---------------------------------------
procedure TExcel.fSetRange(CellFrom,CellTo : OleVariant ; Value : Variant);
begin
  xla.Range[CellFrom,CellTo].Value := value;
end;
//---------------------------------------
function TExcel.fGetRange(CellFrom,CellTo : OleVariant) : Variant;
begin
  result := xla.Range[string(CellFrom),string(CellTo)].Value;
end;
//---------------------------------------
procedure TExcel.SaveAs(filename : string);
begin
  if Assigned(xlw) then begin // а если он не создан?
    xlw.SaveAs(
        filename,
        xlWorkbookNormal,
        '',
        '',
        False,
        False,
        xlNoChange,
        xlLocalSessionChanges,
        true,
        0,
        0,
        LCID);
  end;      
end;
//---------------------------------------

end.
