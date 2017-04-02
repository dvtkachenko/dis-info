unit veksel_isdUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, excel_type;

const
  sAllPage = 'forAllTabSheet';
  sEnterprPage = 'forEnterprTabSheet';
  sSaldoPayVekselPage = 'forSaldoPayVekselTabSheet';
  sChangeVekselPage = 'changeVekselTabSheet';
  sVeksel_isdTemplate = 'veksel_isd.xlt';

type
  TVeksel_isdExportForm = class(TForm)
    Veksel_isdPageControl: TPageControl;
    forAllTabSheet: TTabSheet;
    VekselBeginMaskEdit: TMaskEdit;
    VekselEndMaskEdit: TMaskEdit;
    forEnterprTabSheet: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    forSaldoPayVekselTabSheet: TTabSheet;
    Label5: TLabel;
    changeVekselTabSheet: TTabSheet;
    JournalDateMaskEdit: TMaskEdit;
    Label3: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    allisdVekselQuery: TQuery;
    eventInForVekselQuery: TQuery;
    eventOutForVekselQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure mainVekselReport(Sender: TObject);
    procedure VekselToExcel(Excel : TExcel);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    ReportHeader : string;
    BeginDate : TDateTime;
    EndDate : TDateTime;
    PathToProgram : string;
  end;

implementation

uses shared_type;

{$R *.DFM}

//function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
//function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';

{��������� ���������}

{-------------------}

procedure TVeksel_isdExportForm.FormShow(Sender: TObject);
begin
  VekselBeginMaskEdit.Text := startDate;
  VekselEndMaskEdit.Text := DateToStr(Date);
end;

//---------------------------------------------------------------------
// �������� ��������� ������������ ������ � �������� ��������
// �� ���������� ��� (���� ��� 2000)
//---------------------------------------------------------------------
procedure TVeksel_isdExportForm.mainVekselReport(Sender: TObject);
Var
  temp: lcid;
  Excel : TExcel;
  PathToTemplate : string;
const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  Column = 0;
begin
  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);
  try
    Excel := TExcel.Create;
  except
    raise Exception.Create('���������� ������� OLE - ������');
  end;
  //
  PathToTemplate := PathToProgram + '\Template\' + sVeksel_isdTemplate;
  try
    Excel.AddWorkBook(PathToTemplate);
    Excel.Visible := true;
  except
    raise Exception.Create('���������� ��������� Excel');
  end;
  try
    // ���������� ��������������� ����� ������ � Excel
    VekselToExcel(Excel);
    //
  finally
    Excel.free;
    SetThreadLocale(Temp);
  end;

end;

//---------------------------------------------------------------------
// ����� �������� �������� �� ��������� ������ (���� ��� 2000)
//---------------------------------------------------------------------
procedure TVeksel_isdExportForm.VekselToExcel(Excel : TExcel);
Var
  cell : string;
  cellFrom : string;
  cellTo : string;
  info_row : array[1..24] of Variant;
  i : integer;
  flagEventIn : boolean;
  flagEventOut : boolean;
  rowVeksel : integer;
  rowEventIn : integer;
  rowEventOut : integer;
  countEventIn : integer;
  countEventOut : integer;

  { ����������� ���������� }
  countVeksel : integer ;

  veksel_id : int64;
  // ��������� �������
  veksel_no : string;
  emission_date : TDate;
//  sight_date : TDate;
  emission_place : string;
  nominal_amount : real;
  veksel_maker : string;
  veksel_payer : string;
  // �������� ������� �� �������
  in_type_name : string;
  in_wire_date : TDate;
  in_doc_date : TDate;
  in_wire_amount : real;
  creditor_name : string;
  in_contract_no : string;
  // ��������� ������� �� �������
  out_type_name : string;
  out_wire_date : TDate;
  out_doc_date : TDate;
  out_wire_amount : real;
  debitor_name : string;
  out_contract_no : string;

begin

try
  rowVeksel := 2;
  cell := 'A' + IntToStr(rowVeksel);
  Excel.Cell[cell] := ReportHeader;

  { ��������������  ����������� ���������� }
  countVeksel := 0;
  rowVeksel := 6;

  { ������ � ���� ����������� ����� }
  allisdVekselQuery.Open;

  // ---- ---- ----- ������ ����� �� �������� ----- ----- ----- //
  while not allisdVekselQuery.Eof do begin
    countVeksel := countVeksel + 1;
    flagEventIn := false;
    flagEventOut := false;

    // ----- ------
    Update;
    // ----- ----- ������������ ������ � Excel ------ ------ ------ ------ //
    veksel_id := allisdVekselQuery.fieldbyname('bill_id').asinteger;
    // ��������� �������
    veksel_no := allisdVekselQuery.fieldbyname('bill_no').asstring;
    emission_date := allisdVekselQuery.fieldbyname('emission_date').asdatetime;
//    sight_date := allisdVekselQuery.fieldbyname('sight_date').asdatetime;
    emission_place := allisdVekselQuery.fieldbyname('emission_place').asstring;
    nominal_amount := allisdVekselQuery.fieldbyname('nominal_amount').asfloat;
    veksel_maker := allisdVekselQuery.fieldbyname('bill_maker').asstring;
    veksel_payer := allisdVekselQuery.fieldbyname('bill_payer').asstring;

//    info_row[1] := countVeksel;
//    info_row[2] := veksel_no;
//    info_row[3] := emission_date;
//    info_row[6] := emission_place;
//    info_row[7] := nominal_amount;
//    info_row[9] := veksel_maker;
//    info_row[10] := veksel_payer;

    // ������� ����� ������� �� �������
    cell := 'A' + IntToStr(rowVeksel);
    Excel.Cell[cell] := countVeksel;

    // ������� �� �������
    with eventInForVekselQuery do begin
      Close;
      ParamByName('veksel_id').asinteger := veksel_id;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
      Open;
    end;

    with eventOutForVekselQuery do begin
      Close;
      ParamByName('veksel_id').asinteger := veksel_id;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
      Open;
    end;

    rowEventIn := rowVeksel;
    countEventIn := 0;
    // ---- ---- ----- ������ ����� �� �������� �������� ----- ----- ----- //
    while not eventInForVekselQuery.Eof do begin
      flagEventIn := true;

      in_type_name := eventInForVekselQuery.fieldbyname('type_name').asstring;
      in_wire_date := eventInForVekselQuery.fieldbyname('wire_date').asdatetime;
      in_doc_date := eventInForVekselQuery.fieldbyname('doc_date').asdatetime;
      in_wire_amount := eventInForVekselQuery.fieldbyname('amount').asfloat;
      creditor_name := eventInForVekselQuery.fieldbyname('creditor').asstring;
      in_contract_no := eventInForVekselQuery.fieldbyname('contract').asstring;

      countEventIn := countEventIn + 1;

      // ��������� ��������� ������� ��� �������� ����������
      // ������ � Excel
      info_row[1] := veksel_no;
      info_row[2] := emission_date;
      info_row[3] := emission_place;
      info_row[4] := nominal_amount;
      info_row[5] := veksel_maker;
      info_row[6] := veksel_payer;
      // �������� ������� �� �������
      info_row[8]  := in_type_name;
      info_row[9] := in_wire_date;
      info_row[10] := in_doc_date;
      info_row[11] := in_wire_amount;
      info_row[12] := creditor_name;
      info_row[13] := in_contract_no;

      cellFrom := 'B' + IntToStr(rowEventIn);
      cellTo := 'N' + IntToStr(rowEventIn);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 13 do info_row[i] := unAssigned;

      rowEventIn := rowEventIn + 1;
      eventInForVekselQuery.Next;
    end; // ����� ����� �� �������� ��������
    //
    rowEventOut := rowVeksel;
    countEventOut := 0;
    // ---- ---- ----- ������ ����� �� �������� �������� ----- ----- ----- //
    while not eventOutForVekselQuery.Eof do begin
      flagEventOut := true;

      out_type_name := eventOutForVekselQuery.fieldbyname('type_name').asstring;
      out_wire_date := eventOutForVekselQuery.fieldbyname('wire_date').asdatetime;
      out_doc_date := eventOutForVekselQuery.fieldbyname('doc_date').asdatetime;
      out_wire_amount := eventOutForVekselQuery.fieldbyname('amount').asfloat;
      debitor_name := eventOutForVekselQuery.fieldbyname('debitor').asstring;
      out_contract_no := eventOutForVekselQuery.fieldbyname('contract').asstring;

      countEventOut := countEventOut + 1;

      if (not flagEventIn) or (countEventIn < countEventOut) then begin
        // ��������� ��������� ������� ��� �������� ����������
        // ������ � Excel
        info_row[1] := veksel_no;
        info_row[2] := emission_date;
        info_row[3] := emission_place;
        info_row[4] := nominal_amount;
        info_row[5] := veksel_maker;
        info_row[6] := veksel_payer;

        cellFrom := 'B' + IntToStr(rowEventOut);
        cellTo := 'H' + IntToStr(rowEventOut);

        Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
        for i := 1 to 6 do info_row[i] := unAssigned;
      end;

      // ��������� ������� �� �������
      info_row[1] := out_type_name;
      info_row[2] := out_wire_date;
      info_row[3] := out_doc_date;
      info_row[4] := out_wire_amount;
      info_row[5] := debitor_name;
      info_row[6] := out_contract_no;

      cellFrom := 'P' + IntToStr(rowEventOut);
      cellTo := 'U' + IntToStr(rowEventOut);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 23 do info_row[i] := unAssigned;

      rowEventOut := rowEventOut + 1;
      eventOutForVekselQuery.Next;
    end; // ����� ����� �� �������� ��������
    //
    if (rowEventIn > rowEventOut) then
      rowVeksel := rowEventIn
    else
      rowVeksel := rowEventOut;

    rowVeksel := rowVeksel + 1;
    allisdVekselQuery.Next;
  end; // ����� ����� �� ��������

finally
  allisdVekselQuery.Close;
  eventInForVekselQuery.Close;
  eventOutForVekselQuery.Close;
end;
end;

//---------------------------------------------------------------
procedure TVeksel_isdExportForm.sbReportToExcelClick(Sender: TObject);
//Var
//  id : integer;
//  name : string;
//  s : array[0..maxPChar] of Char;
//  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { ��������������� �������� }
//  pname := @s;
  BeginDate := StrToDate(VekselBeginMaskEdit.Text);
  EndDate := StrToDate(VekselEndMaskEdit.Text);

  if Veksel_isdPageControl.ActivePage.Name = sAllPage then
       begin
         with allisdVekselQuery do begin
           Close;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end;
         ReportHeader := '������ �������� � ���������� ��� �� ������ � ' +
                  VekselBeginMaskEdit.Text + ' �� ' + VekselEndMaskEdit.Text;

         // ��������� �����
         mainVekselReport(Sender);
       end; // ����� iAllPage
{
  if Veksel_isdPageControl.ActivePage.Name = sEnterprPage then
       begin
         with allVekselQuery do begin
           Close;
           SQL.Clear;
           SQL.Add('SELECT o.operation_id,e.enterprise_name creditor,');
           SQL.Add('e1.enterprise_name debitor, O.PAY_DATE,');
           SQL.Add('O.AMOUNTHRIVN,O.AMOUNT_USD, s.type_name, O.COMMENTS , o.contract_no');
           SQL.Add('FROM OPERATIONS O, source_types s,  enterpr e, enterpr e1');
           SQL.Add('WHERE s.type_id = o.type_id');
           SQL.Add('AND (o.creditor_id = e.enterpr_id)');
           SQL.Add('AND (o.debitor_id = e1.enterpr_id)');
           SQL.Add('AND ((o.type_id = 4) or (o.type_id = 5) or (o.type_id = 15) or (o.type_id = 20))');
           SQL.Add('AND o.pay_date >= :begin_date');
           SQL.Add('AND o.pay_date <= :end_date');
           SQL.Add('AND ((o.creditor_id = :ent_id) or (o.debitor_id = :ent_id))');
           SQL.Add('ORDER BY o.creditor_id, O.PAY_DATE,O.AMOUNTHRIVN, O.COMMENTS');
           Prepare;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end;
{
         if GetEnterprise(id,pname) = mrOk then begin
           name := string(pname);
           allVekselQuery.ParamByName('ent_id').asinteger := id;
         end
         else
          raise Exception.Create('����������� �� �������');
}
{         ReportHeader := '������� �� ������ � ' +
                VekselBeginMaskEdit.Text + ' �� ' + VekselEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
         // ��������� �����
         ExportVeksel(Sender);
       end; // ����� sEnterprPage

  if Veksel_isdPageControl.ActivePage.Name = sSaldoPayVekselPage then
       begin
         ReportHeader := '������ �� �������� ������������� � ������ � ��������� ����� ��������  �� '
                         + VekselEndMaskEdit.Text +
                         '    (' + TimeToStr(Time) + ')';
         // ��������� �����
//         ExportVekselSaldo(Sender);
       end; // ����� sSaldoPayVekselPage
}
  Application.BringToFront;
  MessageDlg('������� � Excel ��������', mtInformation, [mbOk], 0);
end;

procedure TVeksel_isdExportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
