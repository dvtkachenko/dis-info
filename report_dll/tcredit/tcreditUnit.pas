unit tcreditUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin;

const
  sreportSPPage = 'reportSPTabSheet';
  sreportTCChangePage = 'changeTabSheet';
  sreportSPCreditTemplate = 'tovcredit.xlt';
  sChangeTCreditTemplate = 'change_tcredit.xlt';
type
  TTovarCreditForm = class(TForm)
    TovarCreditPageControl: TPageControl;
    allTovarCreditContractQuery: TQuery;
    reportSPTabSheet: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    sum_in_oper_contractQuery: TQuery;
    sum_out_oper_contractQuery: TQuery;
    allContractCheckBox: TCheckBox;
    out_tovar_invQuery: TQuery;
    out_interest_invQuery: TQuery;
    in_veksel_payingQuery: TQuery;
    in_money_payingQuery: TQuery;
    contract_saldoQuery: TQuery;
    in_other_payingQuery: TQuery;
    in_invoiceQuery: TQuery;
    check_oper_contractQuery: TQuery;
    check_other_outQuery: TQuery;
    out_otherQuery: TQuery;
    contractTCTabSheet: TTabSheet;
    reportSPBeginMaskEdit: TMaskEdit;
    reportSPEndMaskEdit: TMaskEdit;
    contractTKDBGrid: TDBGrid;
    contractTCDataSource: TDataSource;
    contractAddBitBtn: TBitBtn;
    contractDelBitBtn: TBitBtn;
    insertContractTCQuery: TQuery;
    deleteContractTCQuery: TQuery;
    changeTabSheet: TTabSheet;
    Label3: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    journalDateMaskEdit: TMaskEdit;
    Label9: TLabel;
    changeBeginMaskEdit: TMaskEdit;
    changeEndMaskEdit: TMaskEdit;
    changeTCQuery: TQuery;
    trade_mark_invQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure exportReportSPCredit(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure contractTCTabSheetExit(Sender: TObject);
    procedure contractAddBitBtnClick(Sender: TObject);
    procedure contractDelBitBtnClick(Sender: TObject);
    procedure contractTCTabSheetShow(Sender: TObject);
    procedure exportChangeTCredit(Sender: TObject);
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

uses shared_type, excel_type;

{$R *.DFM}

function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';
function GetContract(id:integer;Var contract_id:integer;Var pcontract_no: PChar) : integer; external 'service.dll' name 'GetContract';


{сервисные процедуры}

{-------------------}

procedure TTovarCreditForm.FormShow(Sender: TObject);
begin
  reportSPBeginMaskEdit.Text := startDate;
  reportSPEndMaskEdit.Text := DateToStr(Date);
  changeBeginMaskEdit.Text := startDate;
  changeEndMaskEdit.Text := DateToStr(Date);
  journalDateMaskEdit.Text := DateToStr(Date-1);
end;

//---------------------------------------------------------------------
// формирует отчет работе ДП ДИС по договорам поставки
// оборудования на условиях товарного кредита
//---------------------------------------------------------------------
procedure TTovarCreditForm.exportReportSPCredit(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..10] of Variant;
     PathToTemplate : string;
     i : integer;
     row : integer;
     rowTovar : integer;
     rowInterest : integer;
     rowPaying : integer;
     rowTop, rowBottom : integer;

     { контрольные переменные }
     countContract : integer ;

     ent_id : real;
     enterprise_name : string;
     contract_no : string;
     // счета-фактуры на товар
     tovar_pay_date : TDate;
     tovar_inv_date : TDate;
     tovar_inv_amount : real;
     tovar_inv_no : string;
     tovar_short_name : string;
     tovar_dept_name :string;
     // счета-фактуры на % по товарному кредиту
     interest_pay_date : TDate;
     interest_inv_date : TDate;
     interest_inv_amount : real;
     interest_inv_no : string;
     interest_short_name : string;
     // погашение товарного кредита
     type_paying : string;
     paying_date : TDate;
     paying_amount : real;
     paying_doc_no : string;
     paying_comment :string;

     // итоговые суммы
     debit, credit : real;
     contract_saldo_begin : real;
     contract_saldo_end : real;
     contract_saldo_end_test : real;
     itog_tovar_inv_amount : real;
     itog_interest_inv_amount : real;
     itog_other_out_amount : real;    // прочие исходящие активы
     itog_veksel_paying_amount : real;
     itog_money_paying_amount : real;
     itog_other_paying_amount : real;
     itog_invoice_amount : real;
     // сумма для проверки были ли операции по договору
     check_oper_contract_amount :  real;
     check_other_out_amount :  real;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     columnEntName = 'B';
     columnContract = 'C';
     columnTovarB = 'D';
     columnTovarInvDate = 'E';
     columnTovarDept = 'I';
     columnTovarE = 'J';
     columnInterestB = 'J';
     columnInterestInvDate = 'K';
     columnInterestE = 'O';
     columnPayingB = 'O';
     columnPayingE = 'T';
     columnBeginSaldo = 'U';
     columnEndSaldo = 'AC';
     columnItogTovarInv = 'V';
     columnItogInterestInv = 'W';
     columnItogOtherOut = 'X';
     columnItogVekselPaying = 'Y';
     columnItogMoneyPaying = 'Z';
     columnItogOtherPaying = 'AA';
     columnItogInvoice = 'AB';
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    try
      Excel := TExcel.Create;
    except
      raise Exception.Create('Невозможно создать OLE - объект');
    end;

    PathToTemplate := PathToProgram + '\Template\' + sreportSPCreditTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     { инициализируем  контрольные переменные }
     countContract := 0;
     row := 6;
     rowTovar := row + 2;
     rowInterest := row + 2;
     rowPaying := row + 2;

     { просим в базе все договора по товарному кредиту }
     allTovarCreditContractQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allTovarCreditContractQuery.Eof do begin
       countContract := countContract + 1;
       ent_id := allTovarCreditContractQuery.fieldbyname('enterpr_id').asfloat;
       enterprise_name := allTovarCreditContractQuery.fieldbyname('enterprise_name').asstring;
       contract_no := allTovarCreditContractQuery.fieldbyname('contract_no').asstring;

       cell := 'A' + IntToStr(row);
       Excel.Cell[cell] := countContract;
       cell := 'B' + IntToStr(row);
       Excel.Cell[cell] := enterprise_name;
       cell := 'H' + IntToStr(row);
       Excel.Cell[cell] := contract_no;
       // делаем заголовки предприятий жирным
       cellFrom := 'A' + IntToStr(row);
       cellTo := 'AB' + IntToStr(row);
       Excel.RangeFontBold(cellFrom, cellTo, 'товарный кредит');

       // инициализация переменных необходимых
       // для того чтобы в Excel удобно было фильтровать
       rowTop := row + 2;
       rowBottom := row + 2;

       // инициализируем итоговые значения
       contract_saldo_begin := 0;
       contract_saldo_end := 0;
       contract_saldo_end_test := 0;
       itog_tovar_inv_amount := 0;
       itog_interest_inv_amount := 0;
       itog_other_out_amount := 0;
       itog_veksel_paying_amount := 0;
       itog_money_paying_amount := 0;
       itog_other_paying_amount := 0;
       itog_invoice_amount := 0;
       check_oper_contract_amount := 0;
       check_other_out_amount := 0;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       with out_tovar_invQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;

       with out_interest_invQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;

       with in_veksel_payingQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;

       with in_money_payingQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;

       with in_other_payingQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;

       with in_invoiceQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;
       // сальдо по договору на начало периода
       with contract_saldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('saldo_date').asdate := BeginDate - 1;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;
       debit := contract_saldoQuery.fieldbyname('debit').asfloat;
       credit := contract_saldoQuery.fieldbyname('credit').asfloat;
       contract_saldo_begin := debit - credit;
       // сальдо по договору на конец периода
       with contract_saldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('saldo_date').asdate := EndDate;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;
       debit := contract_saldoQuery.fieldbyname('debit').asfloat;
       credit := contract_saldoQuery.fieldbyname('credit').asfloat;
       contract_saldo_end_test := debit - credit;

       while not out_tovar_invQuery.Eof do begin
         tovar_pay_date := out_tovar_invQuery.fieldbyname('pay_date').asdatetime;
         tovar_inv_date := out_tovar_invQuery.fieldbyname('invoice_date').asdatetime;
         tovar_inv_amount := out_tovar_invQuery.fieldbyname('amounthrivn').asfloat;
         tovar_inv_no := out_tovar_invQuery.fieldbyname('invoice_no').asstring;
         tovar_short_name := out_tovar_invQuery.fieldbyname('short_trade_mark').asstring;
         tovar_dept_name := out_tovar_invQuery.fieldbyname('dept_name').asstring;

         // выводим в Excel наименование предприятия и договор
         info_row[1] := enterprise_name;
         info_row[2] := contract_no;
         cellFrom := columnEntName + IntToStr(rowTovar);
         cellTo := columnContract + IntToStr(rowTovar);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 2 do info_row[i] := unAssigned;

         // выводим в Excel инф-ю о счетах
         info_row[1] := tovar_pay_date;
         info_row[2] := tovar_inv_date;
         info_row[3] := tovar_inv_amount;
         info_row[4] := tovar_inv_no;
         info_row[5] := tovar_short_name;
         info_row[6] := tovar_dept_name;
         info_row[7] := ' ';

         cellFrom := columnTovarB + IntToStr(rowTovar);
         cellTo := columnTovarE + IntToStr(rowTovar);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 7 do info_row[i] := unAssigned;

         // сравнение даты счета и учетной даты
         // и обозначение их красным если они не совпадают
         if (tovar_pay_date <> tovar_inv_date) then begin
           cellFrom := columnTovarB + IntToStr(rowTovar);
           cellTo := columnTovarInvDate + IntToStr(rowTovar);
           Excel.FillRangeColor(cellFrom, cellTo, 3);
         end;

         itog_tovar_inv_amount := itog_tovar_inv_amount + tovar_inv_amount;

         rowTovar := rowTovar + 1;
         out_tovar_invQuery.Next;
       end;
       // там где выводилась продукция наименование предприятия и договор
       // напротив каждого счета уже выведены
       rowTop := rowTovar;
       rowBottom := rowTovar;

       while not out_interest_invQuery.Eof do begin
         interest_pay_date := out_interest_invQuery.fieldbyname('pay_date').asdatetime;
         interest_inv_date := out_interest_invQuery.fieldbyname('invoice_date').asdatetime;
         interest_inv_amount := out_interest_invQuery.fieldbyname('amounthrivn').asfloat;
         interest_inv_no := out_interest_invQuery.fieldbyname('invoice_no').asstring;
         interest_short_name := out_interest_invQuery.fieldbyname('short_trade_mark').asstring;

         // выводим в Excel инф-ю о выставленных % по тов.кредиту
         info_row[1] := interest_pay_date;
         info_row[2] := interest_inv_date;
         info_row[3] := interest_inv_amount;
         info_row[4] := interest_inv_no;
         info_row[5] := interest_short_name;
         info_row[6] := ' ';

         cellFrom := columnInterestB + IntToStr(rowInterest);
         cellTo := columnInterestE + IntToStr(rowInterest);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 6 do info_row[i] := unAssigned;

         // сравнение даты счета и учетной даты
         // и обозначение их красным если они не совпадают
         if (interest_pay_date <> interest_inv_date) then begin
           cellFrom := columnInterestB + IntToStr(rowInterest);
           cellTo := columnInterestInvDate + IntToStr(rowInterest);
           Excel.FillRangeColor(cellFrom, cellTo, 3);
         end;

         itog_interest_inv_amount := itog_interest_inv_amount + interest_inv_amount;

         rowInterest := rowInterest + 1;

         // для нормальной работы фильтра в Excel
         if rowInterest > rowBottom then
           rowBottom := rowInterest;

         out_interest_invQuery.Next;
       end;

       while not in_veksel_payingQuery.Eof do begin
         type_paying := in_veksel_payingQuery.fieldbyname('type_name').asstring;
         paying_date := in_veksel_payingQuery.fieldbyname('pay_date').asdatetime;
         paying_amount := in_veksel_payingQuery.fieldbyname('amount').asfloat;
         paying_doc_no := '';
         paying_comment := in_veksel_payingQuery.fieldbyname('comment').asstring;

         // выводим в Excel инф-ю о погашении тов.кредита
         info_row[1] := type_paying;
         info_row[2] := paying_date;
         info_row[3] := paying_amount;
         info_row[4] := paying_doc_no;
         info_row[5] := paying_comment;
         info_row[6] := ' ';

         cellFrom := columnPayingB + IntToStr(rowPaying);
         cellTo := columnPayingE + IntToStr(rowPaying);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 6 do info_row[i] := unAssigned;

         itog_veksel_paying_amount := itog_veksel_paying_amount + paying_amount;

         rowPaying := rowPaying + 1;

         // для нормальной работы фильтра в Excel
         if rowPaying > rowBottom then
           rowBottom := rowPaying;

         in_veksel_payingQuery.Next;
       end;

       rowPaying := rowPaying + 1;
       while not in_money_payingQuery.Eof do begin
         type_paying := in_money_payingQuery.fieldbyname('type_name').asstring;
         paying_date := in_money_payingQuery.fieldbyname('doc_date').asdatetime;
         paying_amount := in_money_payingQuery.fieldbyname('amount').asfloat;
         paying_doc_no := in_money_payingQuery.fieldbyname('pay_order').asstring;
         paying_comment := in_money_payingQuery.fieldbyname('comment').asstring;

         // выводим в Excel инф-ю о погашении тов.кредита
         info_row[1] := type_paying;
         info_row[2] := paying_date;
         info_row[3] := paying_amount;
         info_row[4] := paying_doc_no;
         info_row[5] := paying_comment;
         info_row[6] := ' ';

         cellFrom := columnPayingB + IntToStr(rowPaying);
         cellTo := columnPayingE + IntToStr(rowPaying);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 6 do info_row[i] := unAssigned;

         itog_money_paying_amount := itog_money_paying_amount + paying_amount;

         rowPaying := rowPaying + 1;

         // для нормальной работы фильтра в Excel
         if rowPaying > rowBottom then
           rowBottom := rowPaying;

         in_money_payingQuery.Next;
       end;

       rowPaying := rowPaying + 1;
       while not in_other_payingQuery.Eof do begin
         type_paying := in_other_payingQuery.fieldbyname('type_name').asstring;
         paying_date := in_other_payingQuery.fieldbyname('pay_date').asdatetime;
         paying_amount := in_other_payingQuery.fieldbyname('amount').asfloat;
         paying_doc_no := '';
         paying_comment := in_other_payingQuery.fieldbyname('comment').asstring;

         // выводим в Excel инф-ю о погашении тов.кредита
         info_row[1] := type_paying;
         info_row[2] := paying_date;
         info_row[3] := paying_amount;
         info_row[4] := paying_doc_no;
         info_row[5] := paying_comment;
         info_row[6] := ' ';

         cellFrom := columnPayingB + IntToStr(rowPaying);
         cellTo := columnPayingE + IntToStr(rowPaying);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 6 do info_row[i] := unAssigned;

         itog_other_paying_amount := itog_other_paying_amount + paying_amount;

         rowPaying := rowPaying + 1;

         // для нормальной работы фильтра в Excel
         if rowPaying > rowBottom then
           rowBottom := rowPaying;

         in_other_payingQuery.Next;
       end;

       rowPaying := rowPaying + 1;
       while not in_invoiceQuery.Eof do begin
         type_paying := in_invoiceQuery.fieldbyname('type_name').asstring;
         paying_date := in_invoiceQuery.fieldbyname('pay_date').asdatetime;
         paying_amount := in_invoiceQuery.fieldbyname('amount').asfloat;
         paying_doc_no := '';
         paying_comment := in_invoiceQuery.fieldbyname('comment').asstring;

         // выводим в Excel инф-ю о погашении тов.кредита
         info_row[1] := type_paying;
         info_row[2] := paying_date;
         info_row[3] := paying_amount;
         info_row[4] := paying_doc_no;
         info_row[5] := paying_comment;
         info_row[6] := ' ';

         cellFrom := columnPayingB + IntToStr(rowPaying);
         cellTo := columnPayingE + IntToStr(rowPaying);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 6 do info_row[i] := unAssigned;

         itog_invoice_amount := itog_invoice_amount + paying_amount;

         rowPaying := rowPaying + 1;

         // для нормальной работы фильтра в Excel
         if rowPaying > rowBottom then
           rowBottom := rowPaying;

         in_invoiceQuery.Next;
       end;

       // проверка были ли передача прочих активов на предприятие
       // (векселя, деньги , зчт встр. тр. и т.д. и т.п.)
       with check_other_outQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;
       check_other_out_amount := check_other_outQuery.fieldbyname('amount').asfloat;;

       if (round(check_other_out_amount*100) > 0) then begin
         with out_otherQuery do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           ParamByName('contract_no').asstring := contract_no;
           Open;
         end;
         itog_other_out_amount := out_otherQuery.fieldbyname('amount').asfloat;;
       end;

       // вывод итогов
       cell := columnBeginSaldo + IntToStr(row);
       Excel.Cell[cell] := contract_saldo_begin;
       cell := columnItogTovarInv + IntToStr(row);
       Excel.Cell[cell] := itog_tovar_inv_amount;
       cell := columnItogInterestInv + IntToStr(row);
       Excel.Cell[cell] := itog_interest_inv_amount;
       cell := columnItogVekselPaying + IntToStr(row);
       Excel.Cell[cell] := itog_veksel_paying_amount;
       cell := columnItogMoneyPaying + IntToStr(row);
       Excel.Cell[cell] := itog_money_paying_amount;
       cell := columnItogOtherPaying + IntToStr(row);
       Excel.Cell[cell] := itog_other_paying_amount;
       cell := columnItogInvoice + IntToStr(row);
       Excel.Cell[cell] := itog_invoice_amount;

       contract_saldo_end := contract_saldo_begin - itog_tovar_inv_amount -
                             itog_interest_inv_amount -
                             itog_other_out_amount +
                             itog_veksel_paying_amount +
                             itog_money_paying_amount +
                             itog_other_paying_amount +
                             itog_invoice_amount;
       cell := columnEndSaldo + IntToStr(row);
       Excel.Cell[cell] := contract_saldo_end;

       // сравнение итогов и обозначение их красным если они не совпадают
       if (Round(contract_saldo_end*100) <> Round(contract_saldo_end_test*100)) then begin
         cellFrom := columnEndSaldo + IntToStr(row);
         cellTo := columnEndSaldo + IntToStr(row);
         Excel.FillRangeColor(cellFrom, cellTo, 3);
       end;

       // проверка были ли операции по договору
       with check_oper_contractQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('contract_no').asstring := contract_no;
         Open;
       end;
       check_oper_contract_amount := check_oper_contractQuery.fieldbyname('amount').asfloat;;

       // если конечное сальдо по договору  0 и были операции (зеленый)
       if (Round(contract_saldo_end*100) = 0) and
          (Round(check_oper_contract_amount*100) > 0) then begin
         for i := row to rowBottom - 1 do begin
           cellFrom := columnEntName + IntToStr(i);
           cellTo := columnTovarDept + IntToStr(i);
           Excel.FillRangeColor(cellFrom, cellTo, 35);
           //
           cellFrom := columnInterestE + IntToStr(i);
           cellTo := columnEndSaldo + IntToStr(i);
           Excel.FillRangeColor(cellFrom, cellTo, 35);
         end;
       end;

       // если конечное сальдо по договору  0 и не было операций (желтый)
       if (Round(contract_saldo_end*100) = 0) and
          (Round(check_oper_contract_amount*100) = 0) then begin
         for i := row to rowBottom - 1 do begin
           cellFrom := columnEntName + IntToStr(i);
           cellTo := columnEndSaldo + IntToStr(i);
           Excel.FillRangeColor(cellFrom, cellTo, 6);
         end;
       end;

       // если конечное сальдо по договору  -0,01 (красным)
       if (Round(contract_saldo_end*100) = -1) then begin
         for i := row to rowBottom - 1 do begin
           cellFrom := columnEntName + IntToStr(i);
           cellTo := columnTovarDept + IntToStr(i);
           Excel.FillRangeColor(cellFrom, cellTo, 3);
           //
           cellFrom := columnInterestE + IntToStr(i);
           cellTo := columnEndSaldo + IntToStr(i);
           Excel.FillRangeColor(cellFrom, cellTo, 3);
         end;
       end;

       // выводим итоговое значение прочих
       // расположено здесь чтобы перекрасить поля
       if (round(check_other_out_amount*100) > 0) then begin
         cellFrom := columnItogOtherOut + IntToStr(row);
         cellTo := columnItogOtherOut + IntToStr(row);
         Excel.FillRangeColor(cellFrom, cellTo, 33);
         // красим сальдо
         cellFrom := columnEndSaldo + IntToStr(row);
         cellTo := columnEndSaldo + IntToStr(row);
         Excel.FillRangeColor(cellFrom, cellTo, 33);
         // выводим итоговое значение исходящих прочих
         cell := columnItogOtherOut + IntToStr(row);
         Excel.Cell[cell] := itog_other_out_amount;
       end;

       // обеспечиваем условия для нормальной работы фильтра
       if rowTop < rowBottom then begin
         info_row[1] := enterprise_name;
         info_row[2] := contract_no;
         info_row[3] := ' ';
         for i := rowTop to rowBottom - 1 do begin
           cellFrom := columnEntName + IntToStr(i);
           cellTo := columnTovarB + IntToStr(i);
           Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         end;
         for i := 1 to 3 do info_row[i] := unAssigned;
       end;

       // делаем заголовки предприятий жирным
       cellFrom := 'A' + IntToStr(row);
       cellTo := columnEndSaldo + IntToStr(row);
       Excel.RangeFontBold(cellFrom, cellTo, 'товарный кредит');

       // готовимся к следующему договору
       row := rowBottom ;

       // подчеркиваем
       cellFrom := 'A' + IntToStr(row);
       cellTo := columnEndSaldo + IntToStr(row);
       Excel.BottomBordersLine(cellFrom, cellTo, 'товарный кредит');

       row := row + 3;
       rowTovar := row + 2;
       rowInterest := row + 2;
       rowPaying := row + 2;

       allTovarCreditContractQuery.Next;
     end;

   finally
     Excel.free;
     allTovarCreditContractQuery.Close;
     sum_in_oper_contractQuery.Close;
     sum_out_oper_contractQuery.Close;
     out_tovar_invQuery.Close;
     out_interest_invQuery.Close;
     in_veksel_payingQuery.Close;
     in_money_payingQuery.Close;
     in_other_payingQuery.Close;
     in_invoiceQuery.Close;
     check_oper_contractQuery.Close;
     check_other_outQuery.Close;
     out_otherQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

//---------------------------------------------------------------------
// формирование отчета об изменениях в базе за отчетный период
// на определенную дату по договорам товарного кредита
//---------------------------------------------------------------------
procedure TTovarCreditForm.exportChangeTCredit(Sender: TObject);
Var
  temp: lcid;
  Excel : TExcel;
  cell : string;
  cellFrom : string;
  cellTo : string;
  info_row : array[1..16] of Variant;
  PathToTemplate : string;
  i : integer;
  row : integer;
  journalDate : TDate;
  //
  countChange : integer ;
  //
  type_journal : integer;
  type_name_journal : string;
  user_name : string;
  j_pay_date : TDate;
  short_trade_mark : string;
  j_amount : real;
  j_contract_no : string;
  journal_date : TDate;
  debitor_name : string;
  creditor_name : string;
  type_name : string;
  amount : real;
  pay_date : TDate;
  contract_no : string;
  comment : string;

const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
begin
  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  PathToTemplate := PathToProgram + '\Template\' + sChangeTCreditTemplate;
  try
    Excel.AddWorkBook(PathToTemplate);
    Excel.Visible := true;
  except
    raise Exception.Create('Невозможно загрузить Excel');
  end;

  try
    row := 2;
    cell := 'A' + IntToStr(row);
    Excel.Cell[cell] := ReportHeader;

    row := 6;
    JournalDate := StrToDate(JournalDateMaskEdit.Text);

    with changeTCQuery do begin
      Close;
      ParamByName('begin_date').asdate := StrToDate(changeBeginMaskEdit.Text);
      ParamByName('end_date').asdate := StrToDate(changeEndMaskEdit.Text);
      ParamByName('journal_date').asdate := journalDate;
    end;
    changeTCQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
    while not changeTCQuery.Eof do begin
      countChange := countChange + 1;

    // ----- ------
      Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
      type_journal := changeTCQuery.fieldbyname('type_operation').asinteger;
      case type_journal of
        1 : type_name_journal := 'удаление';
        2 : type_name_journal := 'изменение';
        3 : type_name_journal := 'добавление';
      end;
      user_name := changeTCQuery.fieldbyname('user_name').asstring;
      j_pay_date := changeTCQuery.fieldbyname('o_pay_date').asdatetime;
      short_trade_mark := changeTCQuery.fieldbyname('short_trade_mark').asstring;
      j_amount := changeTCQuery.fieldbyname('o_summa').asfloat;
      j_contract_no := changeTCQuery.fieldbyname('o_contract_no').asstring;
      journal_date := changeTCQuery.fieldbyname('journal_date').asdatetime;
      debitor_name := changeTCQuery.fieldbyname('debitor').asstring;
      creditor_name := changeTCQuery.fieldbyname('creditor').asstring;
      type_name := changeTCQuery.fieldbyname('type_name').asstring;
      amount := changeTCQuery.fieldbyname('amount').asfloat;
      pay_date := changeTCQuery.fieldbyname('pay_date').asdatetime;
      contract_no := changeTCQuery.fieldbyname('contract_no').asstring;
      comment := changeTCQuery.fieldbyname('comments').asstring;

      info_row[1] := type_name_journal;
      info_row[2] := user_name;
      info_row[3] := j_pay_date;
      info_row[4] := short_trade_mark;
      info_row[5] := j_amount;
      info_row[6] := j_contract_no;
      info_row[7] := journal_date;
      info_row[8] := ' ';
      info_row[9] := debitor_name;
      info_row[10] := creditor_name;
      info_row[11] := type_name;
      info_row[12] := amount;
      info_row[13] := pay_date;
      info_row[14] := contract_no;
      info_row[15] := comment;

      cellFrom := 'A' + IntToStr(row);
      cellTo := 'O' + IntToStr(row);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 16 do info_row[i] := unAssigned;

      row := row + 1;
      changeTCQuery.Next;
    end;

  finally
    Excel.free;
    changeTCQuery.Close;
    SetThreadLocale(Temp);
  end;
end;

//---------------------------------------------------------------
procedure TTovarCreditForm.sbReportToExcelClick(Sender: TObject);
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
  BeginDate := StrToDate(reportSPBeginMaskEdit.Text);
  EndDate := StrToDate(reportSPEndMaskEdit.Text);

  if TovarCreditPageControl.ActivePage.Name = sreportSPPage then
       begin
//         with allVekselQuery do begin
//           Close;
//           Prepare;
//           ParamByName('begin_date').asdate := BeginDate;
//           ParamByName('end_date').asdate := EndDate;
//         end;
         ReportHeader := 'Отчет по договорам поставок ' +
                         'на условиях товарного кредита за период с ' +
                  reportSPBeginMaskEdit.Text + ' по ' + reportSPEndMaskEdit.Text
                  + ' (' + TimeToStr(Time) + ')';

         // формируем отчет
         exportReportSPCredit(Sender);
       end; // конец sreportSPPage

  if TovarCreditPageControl.ActivePage.Name = sreportTCChangePage then
       begin
         ReportHeader := 'Журнал изменений в базе данных ДИСа '
                      + 'по договорам товарного кредита'
                      + ' за отчетный период с '
                      + changeBeginMaskEdit.Text
                      + ' по '
                      + changeEndMaskEdit.Text
                      + ' начиная с '
                      + JournalDateMaskEdit.Text
                      + ' (' + TimeToStr(Time) + ')';
         // формируем отчет
         exportChangeTCredit(Sender);
       end; // конец sreportTCChangePage

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TTovarCreditForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TTovarCreditForm.contractTCTabSheetExit(Sender: TObject);
begin
  allTovarCreditContractQuery.Close;
end;

procedure TTovarCreditForm.contractAddBitBtnClick(Sender: TObject);
Var
  id : integer;
  contract_id : integer;
  pname : PChar;
  pcontract_no : PChar;
  s : array[0..maxPChar] of Char;
begin
  pname := @s;
  pcontract_no := @s;

  if GetEnterprise(id,pname) = mrOk then begin
    if GetContract(id,contract_id,pcontract_no) = mrOk then begin
      with insertContractTCQuery do begin
        Close;
        ParamByName('contract_id').asinteger:= contract_id;
      end;
      insertContractTCQuery.ExecSQL;
      allTovarCreditContractQuery.Close;
      allTovarCreditContractQuery.Open;
    end
    else
      raise Exception.Create('Договор не выбран');
  end
  else
    raise Exception.Create('Предприятие не выбрано');
end;

procedure TTovarCreditForm.contractDelBitBtnClick(Sender: TObject);
Var
  contract_id : integer;
  contract_no : string;
begin
  contract_no := allTovarCreditContractQuery.fieldbyname('contract_no').asstring;
  if MessageDlg('Вы действительно хотите удалить договор ' + contract_no + ' ?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    contract_id := allTovarCreditContractQuery.fieldbyname('contract_id').asinteger;
    with deleteContractTCQuery do begin
      Close;
      ParamByName('contract_id').asinteger:= contract_id;
    end;
    deleteContractTCQuery.ExecSQL;
    // перечитываем данные 
    allTovarCreditContractQuery.Close;
    allTovarCreditContractQuery.Open;
  end;
end;

procedure TTovarCreditForm.contractTCTabSheetShow(Sender: TObject);
begin
  allTovarCreditContractQuery.Open;
end;

end.
