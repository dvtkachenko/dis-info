unit StatisticReportUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Menus, DBTables, service_type, ActnList,
  ComCtrls, xDBTree, Db, Buttons, ToolWin, ImgList, Mask, ComObj;

const
  sStatisticTemplate = 'statistic.xls';
  sStatisticPlategiTemplate = 'plategi.xls';
  sStatisticDvigenieTemplate = 'dvigenie.xls';
  sStatisticPageName = 'StatisticTabSheet';
  sCoalSenderPageName = 'CoalSenderTabSheet';
  sPlategiPageName = 'allPlategiTabSheet';

type
  TStatisticReportForm = class(TForm)
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    SpeedButton3: TSpeedButton;
    ToolButton1: TToolButton;
    StatisticPageControl: TPageControl;
    StatisticTabSheet: TTabSheet;
    CoalSenderTabSheet: TTabSheet;
    allPlategiTabSheet: TTabSheet;
    CoalQuery: TQuery;
    disInvoiceOutQuery: TQuery;
    disInvoiceInQuery: TQuery;
    disStatisticQueryCreditor: TQuery;
    disStatisticQueryDebitor: TQuery;
    allCoalSenderQuery: TQuery;
    allSaldoQuery: TQuery;
    disPlategiQueryDebitor: TQuery;
    disPlategiQueryCreditor: TQuery;
    disAnyQueryCreditor: TQuery;
    disAnyQueryDebitor: TQuery;
    allContragentQuery: TQuery;
    allPlategiQuery: TQuery;
    allCoalTestQuery: TQuery;
    PlatBeginMaskEdit: TMaskEdit;
    Label3: TLabel;
    Label4: TLabel;
    PlatEndMaskEdit: TMaskEdit;
    ChangeBeginLabel: TLabel;
    BeginMaskEdit: TMaskEdit;
    ChangeEndLabel: TLabel;
    EndMaskEdit: TMaskEdit;
    AllCoalCheckBox: TCheckBox;
    CoalStatisticProgressBar: TProgressBar;
    Label1: TLabel;
    StatBeginMaskEdit: TMaskEdit;
    Label2: TLabel;
    StatEndMaskEdit: TMaskEdit;
    StatisticCheckBox: TCheckBox;
    ProgressBar1: TProgressBar;
    procedure FormShow(Sender: TObject);
    procedure ExportCoalStatistic;
    procedure ExportStatistic;
    procedure ExportPlategi;
    procedure sbReportToExcelClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    PathToProgram : string;
  end;

var
  StatisticReportForm: TStatisticReportForm;

implementation

uses serviceDataUnit;

{$R *.DFM}

function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';

{сервисные процедуры}

{-------------------}

procedure TStatisticReportForm.FormShow(Sender: TObject);
begin
  BeginMaskEdit.Text := '01.01.2000';
  EndMaskEdit.Text := DateToStr(Date);
  StatBeginMaskEdit.Text := '01.01.2000';
  StatEndMaskEdit.Text := DateToStr(Date);
  PlatBeginMaskEdit.Text := '01.01.2000';
  PlatEndMaskEdit.Text := DateToStr(Date);
end;


//---------------------------------------------------------------------
// формирование статистики по всем поставщикам угля за указываемый период
//---------------------------------------------------------------------
procedure TStatisticReportForm.ExportCoalStatistic;
  Var
     temp: lcid;
     vExcel : Variant;
     id : integer;
     name : string;
     s : array[0..maxPChar] of Char;
     pname : PChar;

     BeginDate : TDateTime;
     EndDate : TDateTime;
     PathToTemplate : string;
     ReportHeader : string;
     ent_id : real;
     ent_name : string;
     rowDebit,rowCredit,row : integer;
     allCoalSenderQuery_str: string;

     allSaldoBegin, allSaldoEnd : real;
     allDebitAccept, allCreditAccept : real;
     allDebitNoAccept, allCreditNoAccept : real;
     allSaldoAccept, allSaldoNoAccept : real;
     allCoalNds : real;

     { контрольные переменные }
     allCoalAmount, allCoalQnty : real;
     allCoalAmountFromQBefore : real;
     allCoalQntyFromQBefore : real;
     allCoalAmountFromQAfter : real;
     allCoalQntyFromQAfter : real;
     countCoalSender : integer ;

     curInvoice_no : string;
     curPay_date : TDate;
     curInvoice_date : TDate;
     curQnty : real;
     curAmount : real;
     curNDS : real;
     price : real;
     curCargo_sender : string;
     curCargo_receiver : string;
     curCargo_date : TDate;
     curIs_in_oper : string;
     curContract : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     ColumnDebit = 10;
     ColumnName = 9;
     ColumnCredit = 3;
begin
     temp := GetThreadLocale;
     SetThreadLocale(English_Locale);
     pname := @s;

     BeginDate := StrToDate(BeginMaskEdit.Text);
     EndDate := StrToDate(EndMaskEdit.Text);

     { конструирование запросов }

     if (AllCoalCheckBox.Checked) then
       begin
         with allCoalSenderQuery do begin
           Close;
           SQL.Clear;
           allCoalSenderQuery_str := 'select distinct sender_id, enterpr_name ' +
            'from balans_report_input_coal_all(:begin_date, :end_date) ' +
            'order by enterpr_name';
           SQL.Add(allCoalSenderQuery_str);
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end
       end
     else
       begin
         if GetEnterprise(id,pname) = mrOk then begin
           ent_id := id;
           name := string(pname);
           with allCoalSenderQuery do begin
             Close;
             SQL.Clear;
//             allCoalSenderQuery_str := 'select distinct sender_id, enterpr_name ' +
//              'from balans_report_input_coal_all(:begin_date, :end_date) ' +
//              'where sender_id = :ent_id ' +
//              'order by enterpr_name';
             allCoalSenderQuery_str := 'select enterpr_id sender_id, enterprise_name enterpr_name ' +
              'from  enterpr ' +
              'where enterpr_id = :ent_id ' +
              'order by enterprise_name';
             SQL.Add(allCoalSenderQuery_str);
//             ParamByName('begin_date').asdate := BeginDate;
//             ParamByName('end_date').asdate := EndDate;
             ParamByName('ent_id').asfloat := ent_id;
           end
         end
         else
           raise Exception.Create('Предприятие не выбрано');
       end;
     // запрос на все входящие счета кроме угольных
     with disInvoiceInQuery do begin
       Close;
       SQL.Clear;
       allCoalSenderQuery_str :=
       'SELECT distinct IS_IN_OPER, PAY_DATE, INVOICE_DATE, AMOUNT, NDS,'+
       ' INVOICE_NO , invoice_id, short_trade_mark, cargo_date, contract'+
       ' FROM  balans_report_input_part_ent(:ent_id, :begin_date, :end_date) I,'+
       ' invoice_items I1, supply s, products p'+
       ' where (I1.INVOICE_ID = I.INVOICE_ID)'+
       ' AND (S.SUPPLY_ID = I1.SUPPLY_ID)'+
       ' AND (P.PROD_ID = S.PROD_ID)'+
       ' AND (P.PROD_GROUP_ID <> 12.0)'+
       ' ORDER BY  INVOICE_DATE, INVOICE_NO, AMOUNT';
       SQL.Add(allCoalSenderQuery_str);
     end;

//     try
//     	vExcel := GetActiveOleObject('Excel.Application');
//     except
       try
         vExcel := CreateOleObject('Excel.Application');
       except
         raise Exception.Create('Невозможно загрузить Excel');
       end;
//     end;
     vExcel.Visible := true;

   try
     PathToTemplate := PathToProgram + '\Template\' + sStatisticDvigenieTemplate;
     vExcel.Application.Workbooks.Open(PathToTemplate);
     ReportHeader := 'Статистика работы за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate);
     row := 2;
     vExcel.ActiveSheet.Cells[row, 3].Value := ReportHeader;

     { формируем список всех поставщиков угля в указанный период }
     allCoalSenderQuery.Close;
     allCoalSenderQuery.Open;
     { инициализируем  контрольные переменные }
     allCoalAmount := 0;
     allCoalQnty := 0;
     allCoalAmountFromQBefore := 0;
     allCoalQntyFromQBefore := 0;
     allCoalAmountFromQAfter := 0;
     allCoalQntyFromQAfter := 0;
     countCoalSender := 0;
     allCoalNds := 0;

     row := 6;
     rowCredit := 6;
     rowDebit := 6;

     if (AllCoalCheckBox.Checked) then
       begin
         with allCoalTestQuery do begin
           Close;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         allCoalAmountFromQBefore := allCoalTestQuery.fieldbyname('testAmount').asfloat;
         allCoalQntyFromQBefore := allCoalTestQuery.fieldbyname('testQnty').asfloat;
       end;

  // ---- ---- ----- начало цикла по предприятиям ----- ----- ----- //
     while not allCoalSenderQuery.Eof do
     begin
       ent_id := allCoalSenderQuery.fieldbyname('sender_id').asfloat;
       ent_name := allCoalSenderQuery.fieldbyname('enterpr_name').asstring;
       countCoalSender := countCoalSender + 1;

       with CoalQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;

       with disInvoiceOutQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;

       with disInvoiceInQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;

       with disStatisticQueryDebitor do begin
         Close;
         ParamByName('creditor_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;

       with disStatisticQueryCreditor do begin
         Close;
         ParamByName('debitor_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;

       allSaldoBegin := 0;
       with allSaldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         { берем сальдо на день раньше }
         ParamByName('saldo_date').asdate := BeginDate - 1;
         Open;
       end;
       allSaldoBegin := allSaldoQuery.fieldbyname('allSaldo').asfloat;

       allSaldoEnd := 0;
       with allSaldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('saldo_date').asdate := EndDate;
         Open;
       end;
       allSaldoEnd := allSaldoQuery.fieldbyname('allSaldo').asfloat;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

       allDebitAccept := 0;
       allCreditAccept := 0;
       allDebitNoAccept := 0;
       allCreditNoAccept := 0;
       allSaldoAccept := 0;
       allSaldoNoAccept := 0;

       row := row + 1;
       vExcel.ActiveSheet.Cells[row,3].Value :=
            '-----------------------------------------------------------' +
            '-----------------------------------------------------------';
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnName].Value := ent_name;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 0].Value := 'Сальдо на начало периода:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allSaldoBegin;
       row := row + 2;
       rowDebit := row;
       rowCredit := row;

       { формирование дебитовой статистики}
       { отгрузка угля с предприятия}
       while not CoalQuery.eof do begin
         curInvoice_no := CoalQuery.fieldbyname('invoice_no').asstring;
         curPay_date := CoalQuery.fieldbyname('pay_date').asdatetime;
         curQnty := CoalQuery.fieldbyname('qnty').asfloat;
         curAmount := CoalQuery.fieldbyname('amount').asfloat;
         curNDS := CoalQuery.fieldbyname('nds').asfloat;
         { формируем проверочные значения }
         { должно совпадать с результатами запроса
          на весь уголь отгруженный на ДИС}
         allCoalAmount := allCoalAmount + curAmount;
         allCoalQnty := allCoalQnty + curQnty;
         allCoalNds := allCoalNds + curNDS;

         if curQnty <> 0 then price := curAmount/curQnty;
         curCargo_sender := CoalQuery.fieldbyname('cargo_sender').asstring;
         curCargo_receiver := CoalQuery.fieldbyname('cargo_receiver').asstring;
         curCargo_date := CoalQuery.fieldbyname('cargo_date').asdatetime;
         curIs_in_oper := CoalQuery.fieldbyname('is_in_oper').asstring;
         curContract := CoalQuery.fieldbyname('contract').asstring;
         if curIs_in_oper = 'Y' then allDebitAccept := allDebitAccept + curAmount;
         allDebitNoAccept := allDebitNoAccept + curAmount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := curInvoice_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit+ 1].Value := curPay_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := CoalQuery.fieldbyname('trade_mark').asstring;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := CoalQuery.fieldbyname('dimention').asstring;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := curQnty;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := curAmount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := price;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := curCargo_sender;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := curCargo_receiver;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := curCargo_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 10].Value := curNDS;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value := curIs_in_oper;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 12].Value := curContract;

         rowDebit := rowDebit + 1;
         CoalQuery.Next;
       end;
       CoalQuery.Close;
       rowDebit := rowDebit + 2;

       { отгрузка с предприятия}
       while not disInvoiceInQuery.eof do begin
         curInvoice_no := disInvoiceInQuery.fieldbyname('invoice_no').asstring;
         curPay_date := disInvoiceInQuery.fieldbyname('pay_date').asdatetime;
         curAmount := disInvoiceInQuery.fieldbyname('amount').asfloat;
         curNDS := disInvoiceInQuery.fieldbyname('nds').asfloat;
         curIs_in_oper := disInvoiceInQuery.fieldbyname('is_in_oper').asstring;
         curContract := disInvoiceInQuery.fieldbyname('contract').asstring;
         if curIs_in_oper = 'Y' then allDebitAccept := allDebitAccept + curAmount;
         allDebitNoAccept := allDebitNoAccept + curAmount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := curInvoice_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := curPay_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value :=
                disInvoiceInQuery.fieldbyname('short_trade_mark').asstring;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := curAmount - curNDS;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 10].Value := curNDS;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value :=
                disInvoiceInQuery.fieldbyname('is_in_oper').asstring;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 12].Value := curContract;

         rowDebit := rowDebit + 1;
         disInvoiceInQuery.Next;
       end;
       disInvoiceInQuery.Close;
       rowDebit := rowDebit + 2;

       while not disStatisticQueryDebitor.Eof do begin
         curAmount := disStatisticQueryDebitor.fieldbyname('amounthrivn').asfloat;
         curContract := disStatisticQueryDebitor.fieldbyname('contract').asstring;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := disStatisticQueryDebitor.fieldbyname('pay_date').asdatetime;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := disStatisticQueryDebitor.fieldbyname('type_name').asstring;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := curAmount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 12].Value := curContract;
         allDebitNoAccept := allDebitNoAccept + curAmount;
         allDebitAccept := allDebitAccept + curAmount;

         rowDebit := rowDebit + 1;
         disStatisticQueryDebitor.Next;
       end;
       disStatisticQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       { кредитовая статистика}
       { товарные отгрузки на предприятие }
       while not disInvoiceOutQuery.Eof do begin
         curAmount := disInvoiceOutQuery.fieldbyname('amount').asfloat;
         curNDS := disInvoiceOutQuery.fieldbyname('nds').asfloat;
         curIs_in_oper := disInvoiceOutQuery.fieldbyname('is_in_oper').asstring;
         curContract := disInvoiceOutQuery.fieldbyname('contract').asstring;
         if curIs_in_oper = 'Y' then allCreditAccept := allCreditAccept + curAmount;
         allCreditNoAccept := allCreditNoAccept + curAmount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit - 2].Value := disInvoiceOutQuery.fieldbyname('is_in_oper').asstring;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit - 1].Value := curContract;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := disInvoiceOutQuery.fieldbyname('pay_date').asdatetime;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 2].Value := curAmount - curNDS;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := curNDS;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := disInvoiceOutQuery.fieldbyname('invoice_no').asstring;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 5].Value := disInvoiceOutQuery.fieldbyname('short_trade_mark').asstring;

         rowCredit := rowCredit + 1;
         disInvoiceOutQuery.Next;
       end;
       disInvoiceOutQuery.Close;
       rowCredit := rowCredit + 2;

       { вся кредитовая статистика кроме тов.отгрузок }
       while not disStatisticQueryCreditor.Eof do begin
         curAmount := disStatisticQueryCreditor.fieldbyname('amounthrivn').asfloat;;
         allCreditAccept := allCreditAccept + curAmount;
         allCreditNoAccept := allCreditNoAccept + curAmount;
         curContract := disStatisticQueryCreditor.fieldbyname('contract').asstring;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit - 1].Value := curContract;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := disStatisticQueryCreditor.fieldbyname('pay_date').asdatetime;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 1].Value := curAmount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 5].Value := disStatisticQueryCreditor.fieldbyname('type_name').asstring;

         rowCredit := rowCredit + 1;
         disStatisticQueryCreditor.Next;
       end;
       disStatisticQueryCreditor.Close;
       rowCredit := rowCredit + 2;

       { считаем сальдо  }
       allSaldoNoAccept := allSaldoBegin + allDebitNoAccept - allCreditNoAccept;
       allSaldoAccept := allSaldoBegin + allDebitAccept - allCreditAccept;

       if rowDebit > rowCredit then
         row := rowDebit
       else
         row := rowCredit;

       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 0].Value := 'Полное сальдо на конец периода:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allSaldoEnd;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 0].Value := 'Полное сальдо c акцептом:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allSaldoAccept;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 0].Value := 'Полное сальдо без акцепта:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allSaldoNoAccept;
       row := row + 1;

       // перемещаем указатель на следующее предприятие
       allCoalSenderQuery.Next;
       Update;
     end;

     { контрольные значения после выполнения запросов}
     { это делается дабы отследить возможные изменения в БД
       в процессе формирования полной статистики }
     if (AllCoalCheckBox.Checked) then
       begin
         with allCoalTestQuery do begin
           Close;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         allCoalAmountFromQAfter := allCoalTestQuery.fieldbyname('testAmount').asfloat;
         allCoalQntyFromQAfter := allCoalTestQuery.fieldbyname('testQnty').asfloat;

         row := row + 2;
         vExcel.ActiveSheet.Cells[row,1].Value :=
                '-----------------------------------------------------------';
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 2].Value :=
                'КОНТРОЛЬНЫЕ ЗНАЧЕНИЯ';

         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(сумма) до :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalAmountFromQBefore;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(сумма) после :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalAmountFromQAfter;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(кол-во) до :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalQntyFromQBefore;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(кол-во) после :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalQntyFromQAfter;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(сумма) по отчету :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalAmount;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(кол-во) по отчету :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalQnty;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Всего предприятий поставщиков углей :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := countCoalSender;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Всего входящего НДС (уголь) :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalNDS;
       end;

   finally
     allSaldoQuery.Close;
     CoalQuery.Close;
     allCoalSenderQuery.Close;
     disStatisticQueryDebitor.Close;
     disStatisticQueryCreditor.Close;
     disInvoiceOutQuery.Close;
     disInvoiceInQuery.Close;
     vExcel := unAssigned;
     SetThreadLocale(Temp);
    end;
end;


//---------------------------------------------------------------------
// формирование статистики по всем предприятиям за указываемый период
//---------------------------------------------------------------------
procedure TStatisticReportForm.ExportStatistic;
  Var
     temp: lcid;
     vExcel : Variant;

     id : integer;
     name : string;
     s : array[0..maxPChar] of Char;
     pname : PChar;

     BeginDate : TDateTime;
     EndDate : TDateTime;
     PathToTemplate : string;
     ReportHeader : string;
     ent_id : real;
     ent_name : string;
     rowDebit,rowCredit,row : integer;
     SQL_str: string;

     allSaldoBegin, allSaldoEnd : real;
     allDebitAccept, allCreditAccept : real;
     allDebitNoAccept, allCreditNoAccept : real;
     allSaldoAccept, allSaldoNoAccept : real;

     { контрольные переменные }
     allDebitAmount : real;
     allDebitAmountFromQBefore : real;
     allDebitAmountFromQAfter : real;
     allCreditAmount : real;
     allCreditAmountFromQBefore : real;
     allCreditAmountFromQAfter : real;
     countContragent : integer ;

     pay_date : TDate;
     invoice_date : TDate;
     cargo_date : TDate;
     doc_type : string;
     doc_no : string;
     short_trade_mark : string;
     amount : real;
     amount_usd : real;
     contract_no : string;
     accept : string;
     dept_name : string;
     comment : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     ColumnDebit = 1;
     ColumnName = 13;
     ColumnCredit = ColumnName + 1;
  begin
     temp := GetThreadLocale;
     SetThreadLocale(English_Locale);
     pname := @s;

     BeginDate := StrToDate(StatBeginMaskEdit.Text);
     EndDate := StrToDate(StatEndMaskEdit.Text);

     { конструирование запросов }

     if (StatisticCheckBox.Checked) then
       begin
         // используем объект TQuery для вытягивания всех
         // предприятий контрагентов за указываемй период
         with allContragentQuery do begin
           Close;
           SQL.Clear;
           SQL.Add('select distinct e.enterpr_id, e.enterprise_name ');
           SQL.Add('from operations o, enterpr e');
           SQL.Add('where (o.pay_date >= :begin_date)');
           SQL.Add('and (o.pay_date <= :end_date)');
           SQL.Add('and (o.debitor_id = e.enterpr_id)');
           SQL.Add('and (o.debitor_id <> 0)');  // исключаем ДИС
           SQL.Add('union');
           SQL.Add('select distinct e.enterpr_id, e.enterprise_name');
           SQL.Add('from operations o, enterpr e');
           SQL.Add('where (o.pay_date >= :begin_date)');
           SQL.Add('and (o.pay_date <= :end_date)');
           SQL.Add('and (o.creditor_id = e.enterpr_id)');
           SQL.Add('and (o.creditor_id <> 0)'); // исключаем ДИС
           SQL.Add('union');
           SQL.Add('select distinct e.enterpr_id, e.enterprise_name');
           SQL.Add('from balans_report_all_invoices(:begin_date, :end_date) b, enterpr e');
           SQL.Add('where is_in_oper = ''N'' and (b.sender_id = e.enterpr_id)');
           SQL.Add('and (b.sender_id <> 0)');   // исключаем ДИС
           SQL.Add('union');
           SQL.Add('select distinct e.enterpr_id, e.enterprise_name');
           SQL.Add('from balans_report_all_invoices(:begin_date, :end_date) b, enterpr e');
           SQL.Add('where is_in_oper = ''N'' and (b.payer_id = e.enterpr_id)');
           SQL.Add('and (b.payer_id <> 0)');    // исключаем ДИС

           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end
       end
     else
       begin
         if GetEnterprise(id,pname) = mrOk then begin
           ent_id := id;
           name := string(pname);
           with allContragentQuery do begin
             Close;
             SQL.Clear;
             SQL_str := 'select enterpr_id, enterprise_name ' +
              'from enterpr ' +
              'where enterpr_id = :ent_id ';
             SQL.Add(SQL_str);
             ParamByName('ent_id').asfloat := ent_id;
           end
         end
         else
           raise Exception.Create('Предприятие не выбрано');
       end;
     // запрос на все входящие счета
     with disInvoiceInQuery do begin
       Close;
       SQL.Clear;
       SQL_str :=
       ' SELECT distinct PAY_DATE, INVOICE_DATE, AMOUNT, AMOUNT_USD, NDS,'+
       ' INVOICE_NO , invoice_id, short_trade_mark,'+
       ' cargo_date, contract, is_in_oper, dept_name'+
       ' FROM  balans_report_input_part_ent(:ent_id, :begin_date, :end_date)'+
       ' ORDER BY  PAY_DATE, INVOICE_NO, AMOUNT';
       SQL.Add(SQL_str);
     end;

//     try
//     	vExcel := GetActiveOleObject('Excel.Application');
//     except
       try
         vExcel := CreateOleObject('Excel.Application');
       except
         raise Exception.Create('Невозможно загрузить Excel');
       end;
//     end;
     vExcel.Visible := true;

   try
     PathToTemplate := PathToProgram + '\Template\' + sStatisticTemplate;
     vExcel.Application.Workbooks.Open(PathToTemplate);
     ReportHeader := 'Статистика работы за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate);
     row := 2;
     vExcel.ActiveSheet.Cells[row, 1].Value := ReportHeader;

     { формируем список всех поставщиков угля в указанный период }
     allContragentQuery.Close;
     allContragentQuery.Open;
     { инициализируем  контрольные переменные }
     allDebitAmount := 0;
     allDebitAmountFromQBefore := 0;
     allDebitAmountFromQAfter := 0;
     allCreditAmount := 0;
     allCreditAmountFromQBefore := 0;
     allCreditAmountFromQAfter := 0;
     countContragent := 0;

     row := 7;
     rowCredit := 7;
     rowDebit := 7;

//     if (StatisticCheckBox.Checked) then
//       begin
//         with allCoalTestQuery do begin
//           Close;
//           ParamByName('begin_date').asdate := BeginDate;
//           ParamByName('end_date').asdate := EndDate;
//           Open;
//         end;
//         allCoalAmountFromQBefore := allCoalTestQuery.fieldbyname('testAmount').asfloat;
//         allCoalQntyFromQBefore := allCoalTestQuery.fieldbyname('testQnty').asfloat;
//       end;

  // ---- ---- ----- начало цикла по предприятиям ----- ----- ----- //
     while not allContragentQuery.Eof do begin
       ent_id := allContragentQuery.fieldbyname('enterpr_id').asfloat;
       ent_name := allContragentQuery.fieldbyname('enterprise_name').asstring;
       countContragent := countContragent + 1;

       with disInvoiceOutQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;
       Update;

       with disInvoiceInQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;
       Update;

       with disPlategiQueryDebitor do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;
       Update;

       with disPlategiQueryCreditor do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;
       Update;

       with disAnyQueryDebitor do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;
       Update;

       with disAnyQueryCreditor do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
         Open;
       end;
       Update;

       allSaldoBegin := 0;
       with allSaldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         { берем сальдо на день раньше }
         ParamByName('saldo_date').asdate := BeginDate - 1;
         Open;
       end;
       allSaldoBegin := allSaldoQuery.fieldbyname('allSaldo').asfloat;

       allSaldoEnd := 0;
       with allSaldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('saldo_date').asdate := EndDate;
         Open;
       end;
       allSaldoEnd := allSaldoQuery.fieldbyname('allSaldo').asfloat;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

       allDebitAccept := 0;
       allCreditAccept := 0;
       allDebitNoAccept := 0;
       allCreditNoAccept := 0;
       allSaldoAccept := 0;
       allSaldoNoAccept := 0;

       row := row + 1;
       vExcel.ActiveSheet.Cells[row,3].Value :=
            '-----------------------------------------------------------' +
            '-----------------------------------------------------------';
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnName].Value := ent_name;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'Сальдо на начало периода:';
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := allSaldoBegin;
       row := row + 2;
       rowDebit := row;
       rowCredit := row;

       { формирование дебитовой статистики}

       { отгрузка с предприятия }
       while not disInvoiceInQuery.eof do begin
         pay_date := disInvoiceInQuery.fieldbyname('pay_date').asdatetime;
         invoice_date := disInvoiceInQuery.fieldbyname('invoice_date').asdatetime;
         cargo_date := disInvoiceInQuery.fieldbyname('cargo_date').asdatetime;
         doc_type := 'счет-фактура';
         doc_no := disInvoiceInQuery.fieldbyname('invoice_no').asstring;
         short_trade_mark := disInvoiceInQuery.fieldbyname('short_trade_mark').asstring;
         amount := disInvoiceInQuery.fieldbyname('amount').asfloat;
         amount_usd := disInvoiceInQuery.fieldbyname('amount_usd').asfloat;
         contract_no := disInvoiceInQuery.fieldbyname('contract').asstring;
         accept := disInvoiceInQuery.fieldbyname('is_in_oper').asstring;
         if pay_date < StrToDate('01.01.2000') then
            dept_name := 'unknown'
         else
            dept_name := disInvoiceInQuery.fieldbyname('dept_name').asstring;
//         comment := ;
         if accept = 'Y' then allDebitAccept := allDebitAccept + Amount;
         allDebitNoAccept := allDebitNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := invoice_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := short_trade_mark;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := amount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := accept;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 10].Value := dept_name;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := comment;

         rowDebit := rowDebit + 1;
         disInvoiceInQuery.Next;
       end;
       disInvoiceInQuery.Close;
       rowDebit := rowDebit + 2;

       while not disPlategiQueryDebitor.Eof do begin
         pay_date := disPlategiQueryDebitor.fieldbyname('doc_date').asdatetime;
//         cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disPlategiQueryDebitor.fieldbyname('type_name').asstring;
         doc_no := disPlategiQueryDebitor.fieldbyname('pay_order').asstring;
//         short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disPlategiQueryDebitor.fieldbyname('amount').asfloat;
         amount_usd := disPlategiQueryDebitor.fieldbyname('amount_usd').asfloat;
         contract_no := disPlategiQueryDebitor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disPlategiQueryDebitor.fieldbyname('comment').asstring;
         if accept = 'Y' then allDebitAccept := allDebitAccept + Amount;
         allDebitNoAccept := allDebitNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := doc_no;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := short_trade_mark;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := amount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := accept;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value := comment;

         rowDebit := rowDebit + 1;
         disPlategiQueryDebitor.Next;
       end;
       disPlategiQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       while not disAnyQueryDebitor.Eof do begin
         pay_date := disAnyQueryDebitor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disAnyQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disAnyQueryDebitor.fieldbyname('type_name').asstring;
         doc_no := disAnyQueryDebitor.fieldbyname('act_no').asstring;
//         short_trade_mark := disAnyQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disAnyQueryDebitor.fieldbyname('amount').asfloat;
         amount_usd := disAnyQueryDebitor.fieldbyname('amount_usd').asfloat;
         contract_no := disAnyQueryDebitor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disAnyQueryDebitor.fieldbyname('comment').asstring;
         if accept = 'Y' then allDebitAccept := allDebitAccept + Amount;
         allDebitNoAccept := allDebitNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := doc_no;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := short_trade_mark;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := amount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := accept;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value := comment;

         rowDebit := rowDebit + 1;
         disAnyQueryDebitor.Next;
       end;
       disAnyQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       { кредитовая статистика}
       { товарные отгрузки на предприятие }
       while not disInvoiceOutQuery.Eof do begin
         pay_date := disInvoiceOutQuery.fieldbyname('pay_date').asdatetime;
         invoice_date := disInvoiceOutQuery.fieldbyname('invoice_date').asdatetime;
         cargo_date := disInvoiceOutQuery.fieldbyname('cargo_date').asdatetime;
         doc_type := 'счет-фактура';
         doc_no := disInvoiceOutQuery.fieldbyname('invoice_no').asstring;
         short_trade_mark := disInvoiceOutQuery.fieldbyname('short_trade_mark').asstring;
         amount := disInvoiceOutQuery.fieldbyname('amount').asfloat;
         amount_usd := disInvoiceOutQuery.fieldbyname('amount_usd').asfloat;
         contract_no := disInvoiceOutQuery.fieldbyname('contract').asstring;
         accept := disInvoiceOutQuery.fieldbyname('is_in_oper').asstring;
         if pay_date < StrToDate('01.01.2000') then
            dept_name := 'unknown'
         else
            dept_name := disInvoiceOutQuery.fieldbyname('dept_name').asstring;
//         comment := ;
         if accept = 'Y' then allCreditAccept := allCreditAccept + Amount;
         allCreditNoAccept := allCreditNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 1].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 2].Value := invoice_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 5].Value := short_trade_mark;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := amount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 9].Value := accept;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 10].Value := dept_name;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnCredit + 10].Value := comment;

         rowCredit := rowCredit + 1;
         disInvoiceOutQuery.Next;
       end;
       disInvoiceOutQuery.Close;
       rowCredit := rowCredit + 2;

       { вся кредитовая статистика кроме тов.отгрузок }
       while not disPlategiQueryCreditor.Eof do begin
         pay_date := disPlategiQueryCreditor.fieldbyname('doc_date').asdatetime;
//         cargo_date := disPlategiQueryCreditor.fieldbyname('cargo_date').asdatetime;
         doc_type := disPlategiQueryCreditor.fieldbyname('type_name').asstring;
         doc_no := disPlategiQueryCreditor.fieldbyname('pay_order').asstring;
//         short_trade_mark := disPlategiQueryCreditor.fieldbyname('short_trade_mark').asstring;
         amount := disPlategiQueryCreditor.fieldbyname('amount').asfloat;
         amount_usd := disPlategiQueryCreditor.fieldbyname('amount_usd').asfloat;
         contract_no := disPlategiQueryCreditor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disPlategiQueryCreditor.fieldbyname('comment').asstring;
         if accept = 'Y' then allCreditAccept := allCreditAccept + Amount;
         allCreditNoAccept := allCreditNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := pay_date;
//         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 1].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := doc_no;
//         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := short_trade_mark;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := amount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 9].Value := accept;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 11].Value := comment;

         rowCredit := rowCredit + 1;
         disPlategiQueryCreditor.Next;
       end;
       disPlategiQueryCreditor.Close;
       rowCredit := rowCredit + 2;

       while not disAnyQueryCreditor.Eof do begin
         pay_date := disAnyQueryCreditor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disAnyQueryCreditor.fieldbyname('type_name').asstring;
         doc_no := disAnyQueryCreditor.fieldbyname('act_no').asstring;
//         short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disAnyQueryCreditor.fieldbyname('amount').asfloat;
         amount_usd := disAnyQueryCreditor.fieldbyname('amount_usd').asfloat;
         contract_no := disAnyQueryCreditor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disAnyQueryCreditor.fieldbyname('comment').asstring;
         if accept = 'Y' then allCreditAccept := allCreditAccept + Amount;
         allCreditNoAccept := allCreditNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := pay_date;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnCredit + 1].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := doc_no;
//         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := short_trade_mark;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := amount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 9].Value := accept;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 11].Value := comment;

         rowCredit := rowCredit + 1;
         disAnyQueryCreditor.Next;
       end;
       disAnyQueryCreditor.Close;
       rowCredit := rowCredit + 2;

       { считаем сальдо  }
       allSaldoNoAccept := allSaldoBegin + allDebitNoAccept - allCreditNoAccept;
       allSaldoAccept := allSaldoBegin + allDebitAccept - allCreditAccept;

       if rowDebit > rowCredit then
         row := rowDebit
       else
         row := rowCredit;

       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'Полное сальдо на конец периода:';
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := allSaldoEnd;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'Полное сальдо c акцептом:';
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := allSaldoAccept;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'Полное сальдо без акцепта:';
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := allSaldoNoAccept;
       row := row + 1;

       // перемещаем указатель на следующее предприятие
       allContragentQuery.Next;
       Update;
     end;

     { контрольные значения после выполнения запросов}
     { это делается дабы отследить возможные изменения в БД
       в процессе формирования полной статистики }
{     if (AllCoalCheckBox.Checked) then
       begin
         with allCoalTestQuery do begin
           Close;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         allCoalAmountFromQAfter := allCoalTestQuery.fieldbyname('testAmount').asfloat;
         allCoalQntyFromQAfter := allCoalTestQuery.fieldbyname('testQnty').asfloat;

         row := row + 2;
         vExcel.ActiveSheet.Cells[row,1].Value :=
                '-----------------------------------------------------------';
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 2].Value :=
                'КОНТРОЛЬНЫЕ ЗНАЧЕНИЯ';

         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(сумма) до :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalAmountFromQBefore;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(сумма) после :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalAmountFromQAfter;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(кол-во) до :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalQntyFromQBefore;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(кол-во) после :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalQntyFromQAfter;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(сумма) по отчету :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalAmount;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Уголь(кол-во) по отчету :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := allCoalQnty;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 1].Value := 'Всего предприятий поставщиков углей :';
         vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := countCoalSender;
       end;
}
   finally
     allSaldoQuery.Close;
     allContragentQuery.Close;
     disPlategiQueryDebitor.Close;
     disPlategiQueryCreditor.Close;
     disAnyQueryDebitor.Close;
     disAnyQueryCreditor.Close;
     disInvoiceOutQuery.Close;
     disInvoiceInQuery.Close;
     vExcel := unAssigned;
     SetThreadLocale(Temp);
    end;
end;


//---------------------------------------------------------------------
// формирование отчета по всем платежам за указываемый период
//---------------------------------------------------------------------
procedure TStatisticReportForm.ExportPlategi;
  Var
     temp: lcid;
     vExcel : Variant;
     BeginDate : TDateTime;
     EndDate : TDateTime;
     PathToTemplate : string;
     ReportHeader : string;
     row : integer;

     { контрольные переменные }
//     allDebitAmount : real;
//     allDebitAmountFromQBefore : real;
//     allDebitAmountFromQAfter : real;
//     allCreditAmount : real;
//     allCreditAmountFromQBefore : real;
//     allCreditAmountFromQAfter : real;
     countPlategi : integer ;

     bank_name : string;
     account_num : integer;
     client_name : string;
     client_bank_name : string;
     pay_date : TDateTime;
     debit : real;
     credit : real;
     debit_usd : real;
     credit_usd : real;
     comment : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 1;
begin
     temp := GetThreadLocale;
     SetThreadLocale(English_Locale);

     BeginDate := StrToDate(PlatBeginMaskEdit.Text);
     EndDate := StrToDate(PlatEndMaskEdit.Text);

//     try
//     	vExcel := GetActiveOleObject('Excel.Application');
//     except
       try
         vExcel := CreateOleObject('Excel.Application');
       except
         raise Exception.Create('Невозможно загрузить Excel');
       end;
//     end;
     vExcel.Visible := true;

   try
     PathToTemplate := PathToProgram + '\Template\' + sStatisticPlategiTemplate;
     vExcel.Application.Workbooks.Open(PathToTemplate);
     ReportHeader := 'Все денежные платежи за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate);
     row := 2;
     vExcel.ActiveSheet.Cells[row, 1].Value := ReportHeader;

//     allDebitAmount := 0;
//     allDebitAmountFromQBefore := 0;
//     allDebitAmountFromQAfter := 0;
//     allCreditAmount := 0;
//     allCreditAmountFromQBefore := 0;
//     allCreditAmountFromQAfter := 0;
     countPlategi := 0;

     row := 5;

     with allPlategiQuery do begin
       Close;
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
       Open;
     end;
     Update;

  // ---- ---- ----- начало цикла по всем платежам ----- ----- ----- //
     while not allPlategiQuery.Eof do begin
       countPlategi := countPlategi + 1;

    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       bank_name := allPlategiQuery.fieldbyname('account_name').asstring;
       account_num := allPlategiQuery.fieldbyname('account_id').asinteger;
       client_name := allPlategiQuery.fieldbyname('enterprise_name').asstring;
       client_bank_name := allPlategiQuery.fieldbyname('bank_name').asstring;
       pay_date := allPlategiQuery.fieldbyname('doc_date').asdatetime;
       debit := allPlategiQuery.fieldbyname('debit').asfloat;
       credit := allPlategiQuery.fieldbyname('credit').asfloat;
       debit_usd := allPlategiQuery.fieldbyname('debit_usd').asfloat;
       credit_usd := allPlategiQuery.fieldbyname('credit_usd').asfloat;
       comment := allPlategiQuery.fieldbyname('description').asstring;

       vExcel.ActiveSheet.Cells[row,Column + 0].Value := countPlategi;
       vExcel.ActiveSheet.Cells[row,Column + 1].Value := bank_name;
       vExcel.ActiveSheet.Cells[row,Column + 2].Value := account_num;
       vExcel.ActiveSheet.Cells[row,Column + 3].Value := client_name;
       vExcel.ActiveSheet.Cells[row,Column + 4].Value := client_bank_name;
       vExcel.ActiveSheet.Cells[row,Column + 5].Value := pay_date;
       vExcel.ActiveSheet.Cells[row,Column + 6].Value := debit;
       vExcel.ActiveSheet.Cells[row,Column + 7].Value := credit;
       vExcel.ActiveSheet.Cells[row,Column + 8].Value := debit_usd;
       vExcel.ActiveSheet.Cells[row,Column + 9].Value := credit_usd;
       vExcel.ActiveSheet.Cells[row,Column + 12].Value := comment;

       row := row + 1;
       allPlategiQuery.Next;
     end;

   finally
     Update;
     allPlategiQuery.Close;
     vExcel := unAssigned;
     SetThreadLocale(Temp);
    end;
end;

procedure TStatisticReportForm.sbReportToExcelClick(Sender: TObject);
begin
  //
  if StatisticPageControl.ActivePage.Name = sStatisticPageName then
    ExportStatistic;
  if StatisticPageControl.ActivePage.Name = sCoalSenderPageName then
    ExportCoalStatistic;
  if StatisticPageControl.ActivePage.Name = sPlategiPageName then
    ExportPlategi;
end;

end.
