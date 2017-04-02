unit check_dis_isdUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, excel_type;

// старый кусок запроса для выбрасывания входящего сальдо 
//and d.doc_type_id <> 3100
//and d.doc_type_id <> 3081
//and d.doc_type_id <> 3082
//and d.doc_type_id <> 3101



const
  iComp_isd_disPage = 0;
  sComp_isd_disPage = 'comp_isd_disTabSheet';
  sComp_isd_disTemplate = 'compare_isd_dis.xlt';

type
  Tcomp_dis_isdForm = class(TForm)
    comp_dis_isdPageControl: TPageControl;
    all_contr_oper_isdQuery: TQuery;
    CompBeginMaskEdit: TMaskEdit;
    CompEndMaskEdit: TMaskEdit;
    comp_isd_disTabSheet: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    ruleGroupBox: TGroupBox;
    chainCheckBox: TCheckBox;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    SkidkiPriplCheckBox: TCheckBox;
    saldo_contract_isdQuery: TQuery;
    all_contr_oper_disQuery: TQuery;
    saldo_contract_disQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure comp_dis_isdReport(Sender: TObject);
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

function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';

{сервисные процедуры}

{-------------------}

procedure Tcomp_dis_isdForm.FormShow(Sender: TObject);
begin
  CompBeginMaskEdit.Text := startDate;
  CompEndMaskEdit.Text := DateToStr(Date);
end;


//---------------------------------------------------------------------
// процедура выполняет формирование отчета для сверки
// статистики работы между ИС ДИС98  и  ИС ИСД2000
//---------------------------------------------------------------------
procedure Tcomp_dis_isdForm.comp_dis_isdReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..6] of Variant;
     PathToTemplate : string;
     i, row, row_dis, row_isd : integer;
     cur_contract, prev_contract : string;
     cur_contract_id, prev_contract_id : integer;

     { контрольные переменные }
     contract_debit : real;
     contract_credit : real;
     all_debit : real;
     all_credit : real;
     all_saldo : real;

     // compare variables master
     contract_id : integer;
     contract_no : string;
     signing_date : TDate;

     // compare variables slave
     type_name : string;
     debit : real;
     credit : real;
//     debit_usd : real;
//     credit_usd : real;
     saldo_contract : real;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sComp_isd_disTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try
     ///////////////////////////////////////////////////////////
     // формирование сравнительной статистики
     ///////////////////////////////////////////////////////////
     row := 2;
     cell := 'A' + IntToStr(row);
     ReportHeader := 'Сравнительные данные за период с ' +
                     CompBeginMaskEdit.Text + ' по ' +
                     CompEndMaskEdit.Text;

     Excel.Cell[cell] := ReportHeader;

     // просим в базе ДИС 98 статистику по договорам по которым
     // были операции с Корпорацией ИСД за указанный период
     with all_contr_oper_disQuery do begin
       Close;
       Prepare;
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
     end;
     all_contr_oper_disQuery.Open;

     // просим в базе ИСД2000 статистику по договорам по которым
     // были операции с ДП Корпорации ИСД ДИС за указанный период
     with all_contr_oper_isdQuery do begin
       Close;
       Prepare;
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
     end;
     all_contr_oper_isdQuery.Open;

     { инициализируем  контрольные переменные }
     row := 7;
     contract_debit := 0;
     contract_credit := 0;
     saldo_contract := 0;
     all_debit := 0;
     all_credit := 0;
     all_saldo := 0;

  // ---- ---- ----- начало цикла по договорам  ДИС 98----- ----- ----- //
     row_dis := row;
     // делаем чтобы cur <> prev
     cur_contract := 'butor1';
     prev_contract := 'butor2';
     while not all_contr_oper_disQuery.Eof do begin
       cur_contract := all_contr_oper_disQuery.fieldbyname('contract_no').asstring;
       cur_contract := trim(cur_contract);

       if prev_contract <> cur_contract then begin
         //
         cellFrom := 'A' + IntToStr(row_dis);
         cellTo := 'E' + IntToStr(row_dis);
         Excel.BottomBordersLine(cellFrom,cellTo,'compare');

         row_dis := row_dis + 1;

         contract_no := cur_contract;
         signing_date := all_contr_oper_disQuery.fieldbyname('signing_date').asdatetime;

         with saldo_contract_disQuery do begin
           Close;
           ParamByName('contract_id').asstring := contract_no;
           ParamByName('saldo_date').asdate := EndDate;
         end;
         saldo_contract_disQuery.Open;
         saldo_contract := saldo_contract_disQuery.fieldbyname('saldo_contract').asfloat;

         info_row[1] := contract_no;
         info_row[2] := signing_date;
         info_row[5] := saldo_contract;
         cellFrom := 'A' + IntToStr(row_dis);
         cellTo := 'E' + IntToStr(row_dis);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         Excel.RangeFontBold(cellFrom,cellTo,'compare');

         row_dis := row_dis + 1;
       end;  // конец if prev_contract <> cur_contract

       for i := 1 to 7 do info_row[i] := unAssigned;

       type_name := all_contr_oper_disQuery.fieldbyname('type_name').asstring;
       debit := all_contr_oper_disQuery.fieldbyname('debit').asfloat;
       credit := all_contr_oper_disQuery.fieldbyname('credit').asfloat;

       info_row[1] := type_name;
       info_row[3] := debit;
       info_row[4] := credit;
       cellFrom := 'A' + IntToStr(row_dis);
       cellTo := 'E' + IntToStr(row_dis);
       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 7 do info_row[i] := unAssigned;

       row_dis := row_dis + 1;
       prev_contract := cur_contract;
       all_contr_oper_disQuery.Next;
     end;

  // ---- ---- ----- начало цикла по договорам  ИСД 2000----- ----- ----- //
     row_isd := row;
     // делаем чтобы cur <> prev
     cur_contract_id := 0;
     prev_contract_id := -1;
     while not all_contr_oper_isdQuery.Eof do begin
       cur_contract_id := all_contr_oper_isdQuery.fieldbyname('contract_id').asinteger;

       if prev_contract_id <> cur_contract_id then begin
         //
         cellFrom := 'G' + IntToStr(row_isd);
         cellTo := 'K' + IntToStr(row_isd);
         Excel.BottomBordersLine(cellFrom,cellTo,'compare');

         row_isd := row_isd + 1;

         contract_id := cur_contract_id;
         contract_no := all_contr_oper_isdQuery.fieldbyname('contract_no').asstring;
         signing_date := all_contr_oper_isdQuery.fieldbyname('signing_date').asdatetime;

         with saldo_contract_isdQuery do begin
           Close;
           ParamByName('contract_id').asfloat := contract_id;
           ParamByName('saldo_date').asdate := EndDate;
         end;
         saldo_contract_isdQuery.Open;
         saldo_contract := saldo_contract_isdQuery.fieldbyname('saldo_contract').asfloat;

         info_row[1] := contract_no;
         info_row[2] := signing_date;
         info_row[5] := saldo_contract;
         cellFrom := 'G' + IntToStr(row_isd);
         cellTo := 'K' + IntToStr(row_isd);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         Excel.RangeFontBold(cellFrom,cellTo,'compare');

         row_isd := row_isd + 1;
       end;  // конец if prev_contract_id <> cur_contract_id

       for i := 1 to 7 do info_row[i] := unAssigned;

       type_name := all_contr_oper_isdQuery.fieldbyname('type_name').asstring;
       debit := all_contr_oper_isdQuery.fieldbyname('debit').asfloat;
       credit := all_contr_oper_isdQuery.fieldbyname('credit').asfloat;

       info_row[1] := type_name;
       info_row[3] := debit;
       info_row[4] := credit;
       cellFrom := 'G' + IntToStr(row_isd);
       cellTo := 'K' + IntToStr(row_isd);
       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 7 do info_row[i] := unAssigned;

       row_isd := row_isd + 1;
       prev_contract_id := cur_contract_id;
       all_contr_oper_isdQuery.Next;
     end;

   finally
     Excel.free;
     all_contr_oper_disQuery.Close;
     all_contr_oper_isdQuery.Close;
     saldo_contract_disQuery.Close;
     saldo_contract_isdQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

// ---------------------------------------------------------------
procedure Tcomp_dis_isdForm.sbReportToExcelClick(Sender: TObject);
//Var
//  id : integer;
//  name : string;
//  s : array[0..maxPChar] of Char;
//  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
//  pname := @s;
  BeginDate := StrToDate(CompBeginMaskEdit.Text);
  EndDate := StrToDate(CompEndMaskEdit.Text);

  case comp_dis_isdPageControl.ActivePage.TabIndex of

    iComp_isd_disPage :
       begin
         comp_dis_isdReport(Sender);
       end; // конец iComp_isd_disPage

  end;  // end of CASE

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure Tcomp_dis_isdForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
