unit StatisticReportUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Menus, DBTables, shared_type, ActnList,
  ComCtrls, Db, Buttons, ToolWin, ImgList, Mask, ComObj, excel_type;

const
  sStatisticTemplate = 'statistic.xls';
  sStatisticPlategiTemplate = 'plategi.xls';
  sStatisticDvigenieTemplate = 'dvigenie.xls';
  sStatisticNewTemplate = 'statistic_new.xls';
  sCoalReportTemplate = 'coal_for_balans.xls';  // шаблон для отчета по углю для тряпки
  sStatisticPageName = 'StatisticTabSheet';
  sCoalSenderPageName = 'CoalSenderTabSheet';
  sCoalBalansPageName = 'allCoalBalansTabSheet';
  sPlategiPageName = 'allPlategiTabSheet';
  sall_opPageName = 'all_dis_operTabSheet';

  sStatForBalansPageName = 'StatForBalansTabSheet';
  sCoalReportPageName = 'allCoalReportTabSheet';  // закладка для формирования угля для тряпки
  sall_dis_operation = 'all_dis_operation';  // шаблон для отчета по всем операциям ДИСа
                                             // за указанный период

  // размерность массива для формирования отчета по углю для сводной
  // угольной таблицы
  iMaxItems = 20 ;

  // коды (supply_id) марок углей
  iG = 349;         // марка Г
  iK = 386;         // марка К
  iDG = 403;        // марка ДГ
  iOC = 412;        // марка ОС
  iGG = 413;        // марка Ж
  iDGKOM = 5214;    // марка ДГКОМ
  iT = 5248;        // марка Т
      // рядовые угли
  iAK = 5215;       // марка АК
  iAKO = 867;       // марка АКО
  iAO = 957;        // марка АО
  iAC = 436;        // марка АС
  iGr = 272;        // марка Гр коксующийся
  iGGr = 209;       // марка Жр
  iKr = 501;        // марка Кр
  iOCr = 271;       // марка ОСр
  iTr = 5368;       // марка Тр коксующийся

type
  // сервисные структуры для формирования сводного угольного баланса
  CoalString = record
    coal_name_id : integer;  // соответствует supply_id из базы ДИСа
    coal_name : string;
    qnty : real;
    pure_sum_free_vat : real; // чистая сумма полученная по ф-ле qnty*price
//    pripl_skidki : real;  // приплаты скидки указанные в счете
    add_pripl_skidki : real;  // приплаты-скидки вычисленные пропорциональным методом
    nds : real;
    cargo_receiver : string;
    receiver : string;  //  исп-ся в слчае если данный уголь продается
  end;

  CoalGroupString = record
    Coal : array [1..iMaxItems] of CoalString;
    count : integer;
    all_pripl_skidki : real;
    all_sum_free_vat : real;
    all_nds : real;
  end;

  CoalSender = record
    Group : array [1..iMaxItems] of CoalGroupString;
    count : integer;
  end;

  TStatisticReportForm = class(TForm)
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ExitSpeedButton: TSpeedButton;
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
    StatForBalansTabSheet: TTabSheet;
    Label5: TLabel;
    BeginNewMaskEdit: TMaskEdit;
    Label6: TLabel;
    EndNewMaskEdit: TMaskEdit;
    ExtraInvoiceItemsQuery: TQuery;
    InvoiceItemsQuery: TQuery;
    allInvQuery: TQuery;
    allCoalBalansTabSheet: TTabSheet;
    Label7: TLabel;
    balBeginMaskEdit: TMaskEdit;
    Label8: TLabel;
    balEndMaskEdit: TMaskEdit;
    GetEnterprNameQuery: TQuery;
    CoalOnlyCheckBox: TCheckBox;
    allCoalReportTabSheet: TTabSheet;
    Label9: TLabel;
    BeginCoalReportMaskEdit: TMaskEdit;
    Label10: TLabel;
    EndCoalReportMaskEdit: TMaskEdit;
    notActiveCheckBox: TCheckBox;
    contractCheckBox: TCheckBox;
    contractSaldoQuery: TQuery;
    disContractAnyQueryCreditor: TQuery;
    disContractInvoiceInQuery: TQuery;
    disContractInvoiceOutQuery: TQuery;
    disContractPlategiQueryCreditor: TQuery;
    disContractPlategiQueryDebitor: TQuery;
    disContractAnyQueryDebitor: TQuery;
    chainInvQuery: TQuery;
    coalSenderInvQuery: TQuery;
    detailInvCheckBox: TCheckBox;
    all_dis_in_operQuery: TQuery;
    all_dis_out_operQuery: TQuery;
    all_dis_operTabSheet: TTabSheet;
    Label11: TLabel;
    all_opBeginMaskEdit: TMaskEdit;
    Label12: TLabel;
    all_opEndMaskEdit: TMaskEdit;
    is_coal_enterprQuery: TQuery;
    allContractCheckBox: TCheckBox;
    allEnterprContractQuery: TQuery;
    checkContractOperationQuery: TQuery;
    plategiCheckBox: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure ExportCoalStatistic;
    procedure ExportStatistic;
    procedure ExportAllStatistic(Excel : TExcel; Var ipID : integer);
    procedure ExportContractStatistic;
    procedure CreateCoalSenderReport(Var Coal : CoalSender);
    procedure ExportCoalReport(Sender: TObject);

    procedure ExportPlategi;
    procedure ExportStatisticNew;
//    procedure ExportCoalReport(Sender: TObject);
//    procedure CreateCoalReport;
    procedure sbReportToExcelClick(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure AllCoalCheckBoxClick(Sender: TObject);
    procedure StatisticCheckBoxClick(Sender: TObject);
    procedure contractCheckBoxClick(Sender: TObject);
    procedure FormHide(Sender: TObject);
    // процедуры формирования отчета по всем операциям ДИСа
    // за указанный период
    procedure prepare_report_all_dis_operation;
    procedure export_all_dis_operation(Excel : TExcel);
    procedure allContractCheckBoxClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    parentConfig : p_config;
    PathToProgram : string;
  end;


var
  StatisticReportForm: TStatisticReportForm;

implementation

uses serviceDataUnit, invoicesUnit, Excel_TLB;

{$R *.DFM}

function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';
function GetContract(id:integer;Var contract_id:integer;Var pcontract_no: PChar) : integer; external 'service.dll' name 'GetContract';

{сервисные процедуры}

{-------------------}

procedure TStatisticReportForm.FormShow(Sender: TObject);
Var
  BeginDate, EndDate : TDateTime;
begin
  parentConfig.SharedDll.ReadDate(BeginDate,EndDate);
  StatBeginMaskEdit.Text := DateToStr(BeginDate);
  StatEndMaskEdit.Text := DateToStr(EndDate);
//  StatBeginMaskEdit.Text := startDate;
//  StatEndMaskEdit.Text := DateToStr(Date);
  BeginMaskEdit.Text := startDate;
  EndMaskEdit.Text := DateToStr(Date);
  balBeginMaskEdit.Text := startDate;
  balEndMaskEdit.Text := DateToStr(Date);
  PlatBeginMaskEdit.Text := startDate;
  PlatEndMaskEdit.Text := DateToStr(Date);
  BeginNewMaskEdit.Text := startDate;
  EndNewMaskEdit.Text := DateToStr(Date);
  BeginCoalReportMaskEdit.Text := startDate;
  EndCoalReportMaskEdit.Text := DateToStr(Date);
  all_opBeginMaskEdit.Text := startDate;
  all_opEndMaskEdit.Text := DateToStr(Date);
  CoalOnlyCheckBox.Visible := false;
  CoalOnlyCheckBox.Enabled := false;
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
   if StatisticPageControl.ActivePage.Name = sCoalSenderPageName then
    begin
   //
     if (AllCoalCheckBox.Checked) then
       begin
         with allCoalSenderQuery do begin
           Close;
           SQL.Clear;
           allCoalSenderQuery_str := 'select distinct sender_id, enterpr_name ' +
            'from balans_report_input_coal_all(:begin_date, :end_date) ';
            if (CoalOnlyCheckBox.Checked) then begin
              allCoalSenderQuery_str := allCoalSenderQuery_str +
                                        'where prod_id = 10010 or prod_id = 10011 ';
            end;
           allCoalSenderQuery_str := allCoalSenderQuery_str +
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
    end
   else  //  формируем балансы по предприятиям поставщикам углей
     begin
       with allCoalSenderQuery do begin
         Close;
         SQL.Clear;
         SQL.Add('select * from balans_enterpr');
         if (notActiveCheckBox.Checked) then
           SQL.Add('where active_ = ''N''')
         else
           SQL.Add('where active_ = ''Y''');
         SQL.Add('and is_balance = ''Y''');
         SQL.Add('order by group_id, enterpr_id');
//         DatabaseName := 'ora_dis';
         DatabaseName := 'my_dis_ibdb_cyrr';
         AllCoalCheckBox.Checked := true;
         BeginDate := StrToDate(balBeginMaskEdit.Text);
         EndDate := StrToDate(balEndMaskEdit.Text);
       end

     end;

     // запрос на все входящие счета кроме угольных
     with disInvoiceInQuery do begin
       Close;
       SQL.Clear;
       SQL.Add('SELECT distinct IS_IN_OPER, PAY_DATE, INVOICE_DATE, AMOUNT, NDS,');
       SQL.Add('INVOICE_NO , invoice_id, short_trade_mark, cargo_date, contract');
       SQL.Add('FROM  balans_report_input_part_ent(:ent_id, :begin_date, :end_date) I');
       SQL.Add('where not exists');
       SQL.Add('(select * from invoice_items I1, supply s, products p');
       SQL.Add(' where (I1.INVOICE_ID = I.INVOICE_ID)');
       SQL.Add('AND (S.SUPPLY_ID = I1.SUPPLY_ID)');
       SQL.Add('AND (P.PROD_ID = S.PROD_ID)');
       SQL.Add('AND (P.PROD_GROUP_ID = 12.0))');
       SQL.Add('ORDER BY CONTRACT, INVOICE_DATE, INVOICE_NO, AMOUNT');
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
     if StatisticPageControl.ActivePage.Name = sCoalSenderPageName then
      begin
       ent_id := allCoalSenderQuery.fieldbyname('sender_id').asfloat;
       ent_name := allCoalSenderQuery.fieldbyname('enterpr_name').asstring;
      end
     else
      begin
       ent_id := allCoalSenderQuery.fieldbyname('enterpr_id').asfloat;

       GetEnterprNameQuery.Close;
       GetEnterprNameQuery.ParamByName('ent_id').asfloat := ent_id;
       GetEnterprNameQuery.Open;
       ent_name := GetEnterprNameQuery.fieldbyname('enterprise_name').asstring;
      end;

     countCoalSender := countCoalSender + 1;


       with CoalQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
//         if (CoalOnlyCheckBox.Checked) then
//           ParamByName('cox_coal').asinteger := 10010
//         else
           ParamByName('cox_coal').asinteger := 13;
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
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := curAmount;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 10].Value := curNDS;
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
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 2].Value := curAmount;
//         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := curNDS;
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
  old_lang : lcid;
  Excel : TExcel;
  PathToTemplate : string;
  InvExportForm : TInvExportForm;
  invDLL : TReportFormDLL;
  ipID : integer;
const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
begin
    if (contractCheckBox.Checked or allContractCheckBox.Checked) then begin
      ExportContractStatistic
    end
    else begin

    try
      old_lang := GetThreadLocale;
      SetThreadLocale(English_Locale);

      Excel := TExcel.Create;
      PathToTemplate := PathToProgram + '\Template\' + sStatisticTemplate;
      try
        Excel.AddWorkBook(PathToTemplate);
      except
        raise Exception.Create('Невозможно загрузить Excel');
      end;
      ExportAllStatistic(Excel, ipID);
      // производим детализацию статистики по счетам
      if detailInvCheckBox.Checked = true then begin
        InvExportForm := nil;
        invDLL := parentConfig.LibDLL.GetDLLbyFilename('invoices.dll');
        try
          invDLL.InitServiceExternalCall(pointer(InvExportForm));
          if (InvExportForm = nil) then
            raise Exception.Create('Невозможно инициализировать внешний сервис');
//          InvExportForm.InvPageControl.Pages[0].TabVisible := true;
          InvExportForm.ipExcel := Excel;
          InvExportForm.ipID := ipID;
          InvExportForm.InvBeginMaskEdit.Text := StatBeginMaskEdit.Text;
          InvExportForm.InvEndMaskEdit.Text := StatEndMaskEdit.Text;
          InvExportForm.InvPageControl.Pages[0].TabVisible := true;
          // входящие счета - фактуры
          Excel.SelectWorkSheet('in_inv');
          InvExportForm.InOutRadioGroup.ItemIndex := 0;
          InvExportForm.chainCheckBox.Checked := true;
          InvExportForm.sbReportToExcelClick(nil);
          // исходящие счета - факутры
          Excel.SelectWorkSheet('out_inv');
          InvExportForm.InOutRadioGroup.ItemIndex := 1;
          InvExportForm.chainCheckBox.Checked := true;
          InvExportForm.sbReportToExcelClick(nil);
          Excel.SelectWorkSheet('Статистика');
        finally
          invDLL.DeInitServiceExternalCall;
        end;
      end;
    finally
      Excel.free;
      Excel := nil;
      SetThreadLocale(old_lang);
    end;

    end;
 //
end;

//  формирование полной статистики по предприятию
procedure TStatisticReportForm.ExportAllStatistic(Excel : TExcel; Var ipID : integer);
  Var
     id : integer;
     name : string;
     s : array[0..maxPChar] of Char;
     pname : PChar;
     i : integer;

     cell : string;
     cellFrom : string;
     cellTo : string;
     s_row : string;
     info_row : array[1..12] of Variant;

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
//     ColumnDebit = 1;
     ColumnName = 'M';
//     ColumnCredit = ColumnName + 1;
  begin
     pname := @s;

     BeginDate := StrToDate(StatBeginMaskEdit.Text);
     EndDate := StrToDate(StatEndMaskEdit.Text);

     { конструирование запросов }
  try
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
         Application.BringToFront;
         if GetEnterprise(id,pname) = mrOk then begin
           ent_id := id;
           ipID := id;
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
     // показываем Excel 
     Excel.Visible := true;

     // запрос на все входящие счета
     with disInvoiceInQuery do begin
       Close;
       SQL.Clear;
       SQL_str :=
       ' SELECT distinct PAY_DATE, INVOICE_DATE, AMOUNT, AMOUNT_USD, NDS,'+
       ' INVOICE_NO , invoice_id, short_trade_mark,'+
       ' cargo_date, contract, is_in_oper, dept_name'+
       ' FROM  balans_report_input_part_ent(:ent_id, :begin_date, :end_date)'+
       ' ORDER BY CONTRACT, PAY_DATE, INVOICE_NO, AMOUNT';
       SQL.Add(SQL_str);
     end;

     ReportHeader := 'Статистика работы за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate);
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

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
       cell := 'C' + IntToStr(row);
       Excel.Cell[cell] :=
            '-----------------------------------------------------------' +
            '-----------------------------------------------------------';
       row := row + 1;
       cell := ColumnName + IntToStr(row);
       Excel.Cell[cell] := ent_name;

       row := row + 1;
       // cell = ColumnDebit + 3;
       cell := 'D' + IntToStr(row);
       Excel.Cell[cell] := 'Сальдо на начало периода:';
       // cell = ColumnDebit + 6;
       cell := 'G' + IntToStr(row);
       Excel.Cell[cell] := allSaldoBegin;

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

         //-----------------------------------------------
         // блока экспорта в Excel
         info_row[1] := pay_date;
         info_row[2] := cargo_date;
         info_row[3] := invoice_date;
         info_row[4] := doc_type;
         info_row[5] := doc_no;
         info_row[6] := short_trade_mark;
         info_row[7] := amount;
         info_row[8] := amount_usd;
         info_row[9] := contract_no;
         info_row[10] := accept;
         info_row[11] := dept_name;

         cellFrom := 'A' + IntToStr(rowDebit);
         cellTo := 'L' + IntToStr(rowDebit);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 12 do info_row[i] := unAssigned;
         //-----------------------------------------------

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

         //-----------------------------------------------
         // блока экспорта в Excel
         info_row[1] := pay_date;
//         info_row[2] := cargo_date;
//         info_row[3] := invoice_date;
         info_row[4] := doc_type;
         info_row[5] := doc_no;
//         info_row[6] := short_trade_mark;
         info_row[7] := amount;
         info_row[8] := amount_usd;
         info_row[9] := contract_no;
         info_row[10] := accept;
//         info_row[11] := dept_name;
         info_row[12] := comment;

         cellFrom := 'A' + IntToStr(rowDebit);
         cellTo := 'L' + IntToStr(rowDebit);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 12 do info_row[i] := unAssigned;
         //-----------------------------------------------

         rowDebit := rowDebit + 1;
         disPlategiQueryDebitor.Next;
       end;
       disPlategiQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       while not disAnyQueryDebitor.Eof do begin
         pay_date := disAnyQueryDebitor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disAnyQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disAnyQueryDebitor.fieldbyname('type_name').asstring;
//         doc_no := disAnyQueryDebitor.fieldbyname('act_no').asstring;
//         short_trade_mark := disAnyQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disAnyQueryDebitor.fieldbyname('amount').asfloat;
         amount_usd := disAnyQueryDebitor.fieldbyname('amount_usd').asfloat;
         contract_no := disAnyQueryDebitor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disAnyQueryDebitor.fieldbyname('comment').asstring;
         if accept = 'Y' then allDebitAccept := allDebitAccept + Amount;
         allDebitNoAccept := allDebitNoAccept + Amount;

         //-----------------------------------------------
         // блока экспорта в Excel
         info_row[1] := pay_date;
//         info_row[2] := cargo_date;
//         info_row[3] := invoice_date;
         info_row[4] := doc_type;
//         info_row[5] := doc_no;
//         info_row[6] := short_trade_mark;
         info_row[7] := amount;
         info_row[8] := amount_usd;
         info_row[9] := contract_no;
         info_row[10] := accept;
//         info_row[11] := dept_name;
         info_row[12] := comment;

         cellFrom := 'A' + IntToStr(rowDebit);
         cellTo := 'L' + IntToStr(rowDebit);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 12 do info_row[i] := unAssigned;
         //-----------------------------------------------

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

         //-----------------------------------------------
         // блока экспорта в Excel
         info_row[1] := pay_date;
         info_row[2] := cargo_date;
         info_row[3] := invoice_date;
         info_row[4] := doc_type;
         info_row[5] := doc_no;
         info_row[6] := short_trade_mark;
         info_row[7] := amount;
         info_row[8] := amount_usd;
         info_row[9] := contract_no;
         info_row[10] := accept;
         info_row[11] := dept_name;

         cellFrom := 'N' + IntToStr(rowCredit);
         cellTo := 'Y' + IntToStr(rowCredit);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 12 do info_row[i] := unAssigned;
         //-----------------------------------------------

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

         //-----------------------------------------------
         // блока экспорта в Excel
         info_row[1] := pay_date;
//         info_row[2] := cargo_date;
//         info_row[3] := invoice_date;
         info_row[4] := doc_type;
         info_row[5] := doc_no;
//         info_row[6] := short_trade_mark;
         info_row[7] := amount;
         info_row[8] := amount_usd;
         info_row[9] := contract_no;
         info_row[10] := accept;
//         info_row[11] := dept_name;
         info_row[12] := comment;

         cellFrom := 'N' + IntToStr(rowCredit);
         cellTo := 'Y' + IntToStr(rowCredit);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 12 do info_row[i] := unAssigned;
         //-----------------------------------------------

         rowCredit := rowCredit + 1;
         disPlategiQueryCreditor.Next;
       end;
       disPlategiQueryCreditor.Close;
       rowCredit := rowCredit + 2;

       while not disAnyQueryCreditor.Eof do begin
         pay_date := disAnyQueryCreditor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disAnyQueryCreditor.fieldbyname('type_name').asstring;
//         doc_no := disAnyQueryCreditor.fieldbyname('act_no').asstring;
//         short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disAnyQueryCreditor.fieldbyname('amount').asfloat;
         amount_usd := disAnyQueryCreditor.fieldbyname('amount_usd').asfloat;
         contract_no := disAnyQueryCreditor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disAnyQueryCreditor.fieldbyname('comment').asstring;
         if accept = 'Y' then allCreditAccept := allCreditAccept + Amount;
         allCreditNoAccept := allCreditNoAccept + Amount;

         //-----------------------------------------------
         // блока экспорта в Excel
         info_row[1] := pay_date;
//         info_row[2] := cargo_date;
//         info_row[3] := invoice_date;
         info_row[4] := doc_type;
//         info_row[5] := doc_no;
//         info_row[6] := short_trade_mark;
         info_row[7] := amount;
         info_row[8] := amount_usd;
         info_row[9] := contract_no;
         info_row[10] := accept;
//         info_row[11] := dept_name;
         info_row[12] := comment;

         cellFrom := 'N' + IntToStr(rowCredit);
         cellTo := 'Y' + IntToStr(rowCredit);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 12 do info_row[i] := unAssigned;
         //-----------------------------------------------

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
       // cell = ColumnDebit + 3;
       cell := 'D' + IntToStr(row);
       Excel.Cell[cell] := 'Полное сальдо на конец периода:';
       // cell = ColumnDebit + 6;
       cell := 'G' + IntToStr(row);
       Excel.Cell[cell] := allSaldoEnd;

       row := row + 1;
       // cell = ColumnDebit + 3;
       cell := 'D' + IntToStr(row);
       Excel.Cell[cell] := 'Полное сальдо c акцептом:';
       // cell = ColumnDebit + 6;
       cell := 'G' + IntToStr(row);
       Excel.Cell[cell] := allSaldoAccept;

       row := row + 1;
       // cell = ColumnDebit + 3;
       cell := 'D' + IntToStr(row);
       Excel.Cell[cell] := 'Полное сальдо без акцепта:';
       // cell = ColumnDebit + 6;
       cell := 'G' + IntToStr(row);
       Excel.Cell[cell] := allSaldoNoAccept;
       row := row + 1;

       // перемещаем указатель на следующее предприятие
       allContragentQuery.Next;
       Update;
     end;

   finally
     allSaldoQuery.Close;
     allContragentQuery.Close;
     disPlategiQueryDebitor.Close;
     disPlategiQueryCreditor.Close;
     disAnyQueryDebitor.Close;
     disAnyQueryCreditor.Close;
     disInvoiceOutQuery.Close;
     disInvoiceInQuery.Close;
   end;
end;

//  формирование статистики по предприятию
//  по указанному договору

procedure TStatisticReportForm.ExportContractStatistic;
  Var
     temp: lcid;
     vExcel : Variant;
     RangeLeft, RangeRight : string;

     username : array[0..50] of char;
     p_username : pchar;
     len : cardinal;
     str_username : string;

     id : integer;
     contract_id : integer;
     name : string;
     s : array[0..maxPChar] of Char;
     pname : PChar;
     contract : string;
     signing_date : TDateTime;
     pcontract_no : PChar;

     BeginDate : TDateTime;
     EndDate : TDateTime;
     PathToTemplate : string;
     ReportHeader : string;
     ent_id : real;
     ent_name : string;
     rowDebit,rowCredit,row : integer;
     SQL_str: string;

     contractSaldoBegin, contractSaldoEnd : real;
     contractDebitAccept, contractCreditAccept : real;
     contractDebitNoAccept, contractCreditNoAccept : real;
     contractSaldoAccept, contractSaldoNoAccept : real;

     { контрольные переменные }
     countContragent : integer ;

     all_contractes_saldo : real; // сумма всех сальдо по контрактам
     all_saldo : real;            // полное сальдо 

     debit, credit : real;

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
     pcontract_no := @s;

     BeginDate := StrToDate(StatBeginMaskEdit.Text);
     EndDate := StrToDate(StatEndMaskEdit.Text);

     try
       vExcel := CreateOleObject('Excel.Application');
       PathToTemplate := PathToProgram + '\Template\' + sStatisticTemplate;
       vExcel.Application.Workbooks.Open(PathToTemplate);
     except
       raise Exception.Create('Невозможно загрузить Excel');
     end;

     { конструирование запросов }
     Application.BringToFront;
     // получаем имя пользователя
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
       end;
       if (allContractCheckBox.Checked) then begin
         with allEnterprContractQuery do begin
           Close;
           SQL.Clear;
           SQL.Add('select 1, c.contract_no, c.signing_date from contract c, contract_sides cs');
           SQL.Add('where c.contract_no = cs.contract_no and');
           SQL.Add('cs.enterpr_id = :ent_id');
           // добавляем строку с contract_no = "НЕТ ПРИВЯЗКИ!!!"
           SQL.Add('union');
           SQL.Add('select 2, c1.contract_no, cast(null as date) signing_date');
           SQL.Add('from contract c1');
           SQL.Add('where c1.contract_id = 2106');
           SQL.Add('order by 1,3');
           ParamByName('ent_id').asfloat := ent_id;
         end;
       end
       else begin
         if GetContract(id,contract_id,pcontract_no) = mrOk then begin
           contract := string(pcontract_no);
           with allEnterprContractQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select c.contract_no, c.signing_date from contract c, contract_sides cs');
             SQL.Add('where c.contract_no = cs.contract_no and');
             SQL.Add('cs.enterpr_id = :ent_id and');
             SQL.Add('c.contract_no = :contract');
             ParamByName('ent_id').asfloat := ent_id;
             ParamByName('contract').asstring := contract;
           end;
         end
         else
           raise Exception.Create('Договор не выбран');
       end;
     end
     else
       raise Exception.Create('Предприятие не выбрано');
     // выбираем из базы договора по которым будем вытаскивать статистику
     allEnterprContractQuery.Open;

     vExcel.Visible := true;

   try
     ReportHeader := 'Статистика работы по договору ' + contract +
                     ' за период с ' + datetostr(BeginDate) +
                     ' по ' + datetostr(EndDate);
     row := 2;
     vExcel.ActiveSheet.Cells[row, 1].Value := ReportHeader;

     { формируем список всех поставщиков угля в указанный период }
     allContragentQuery.Close;
     allContragentQuery.Open;
     { инициализируем  контрольные переменные }
     countContragent := 0;

     all_contractes_saldo := 0;
     all_saldo := 0;

     row := 7;
     rowCredit := 7;
     rowDebit := 7;

  // ---- ---- ----- начало цикла по предприятиям ----- ----- ----- //
     while not allContragentQuery.Eof do begin
       ent_id := allContragentQuery.fieldbyname('enterpr_id').asfloat;
       ent_name := allContragentQuery.fieldbyname('enterprise_name').asstring;
       countContragent := countContragent + 1;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnName].Value := ent_name;
       row := row + 1;

       // ---- ---- ----- начало цикла по договорам ----- ----- ----- //
       while not allEnterprContractQuery.Eof do begin
         //
         contract := allEnterprContractQuery.fieldbyname('contract_no').asstring;
         signing_date := allEnterprContractQuery.fieldbyname('signing_date').asdatetime;

         contractSaldoBegin := 0;
         with contractSaldoQuery do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           { берем сальдо на день раньше }
           ParamByName('saldo_date').asdate := BeginDate - 1;
           Open;
         end;
         contractSaldoBegin := ContractSaldoQuery.fieldbyname('contractSaldo').asfloat;

         contractSaldoEnd := 0;
         with contractSaldoQuery do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('saldo_date').asdate := EndDate;
           Open;
         end;
         contractSaldoEnd := contractSaldoQuery.fieldbyname('contractSaldo').asfloat;

         debit := 0;
         credit := 0;
         with checkContractOperationQuery do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         debit := checkContractOperationQuery.fieldbyname('debit').asfloat;
         credit := checkContractOperationQuery.fieldbyname('credit').asfloat;

         // если сальдо на начало периода, на конец периода
         // все обороты за периоды равны 0, то статистика по договору не выводится
         if ((Round(contractSaldoBegin*51.00) = 0) and
             (Round(contractSaldoEnd*51.00) = 0) and
             (Round(debit*51.00) = 0) and
             (Round(credit*51.00) = 0)) then begin

             allEnterprContractQuery.Next;
             continue;
         end;

         with disContractInvoiceOutQuery do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         Update;

         with disContractInvoiceInQuery do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         Update;

         with disContractPlategiQueryDebitor do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         Update;

         with disContractPlategiQueryCreditor do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         Update;

         with disContractAnyQueryDebitor do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         Update;

         with disContractAnyQueryCreditor do begin
           Close;
           ParamByName('ent_id').asfloat := ent_id;
           ParamByName('contract_no').asstring := contract;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
           Open;
         end;
         Update;

       // ----- ------
         Update;
       // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

         contractDebitAccept := 0;
         contractCreditAccept := 0;
         contractDebitNoAccept := 0;
         contractCreditNoAccept := 0;
         contractSaldoAccept := 0;
         contractSaldoNoAccept := 0;

         vExcel.ActiveSheet.Cells[row,3].Value :=
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------';

         RangeLeft := 'A' + IntToStr(row);
         RangeRight := 'Y' + IntToStr(row);
         vExcel.ActiveSheet.Range[OleVariant(RangeLeft),OleVariant(RangeRight)].Interior.ColorIndex := 6;

         row := row + 2;
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'Сальдо на начало периода';
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'по договору ' + contract
                                                                + ' от ' + DateToStr(signing_date) + ' :';
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := contractSaldoBegin;
         row := row + 2;
         rowDebit := row;
         rowCredit := row;

         { формирование дебитовой статистики}

         { отгрузка с предприятия }
         while not disContractInvoiceInQuery.eof do begin
           pay_date := disContractInvoiceInQuery.fieldbyname('pay_date').asdatetime;
           invoice_date := disContractInvoiceInQuery.fieldbyname('invoice_date').asdatetime;
           cargo_date := disContractInvoiceInQuery.fieldbyname('cargo_date').asdatetime;
           doc_type := 'счет-фактура';
           doc_no := disContractInvoiceInQuery.fieldbyname('invoice_no').asstring;
           short_trade_mark := disContractInvoiceInQuery.fieldbyname('short_trade_mark').asstring;
           amount := disContractInvoiceInQuery.fieldbyname('amount').asfloat;
           amount_usd := disContractInvoiceInQuery.fieldbyname('amount_usd').asfloat;
           contract_no := disContractInvoiceInQuery.fieldbyname('contract').asstring;
           accept := disContractInvoiceInQuery.fieldbyname('is_in_oper').asstring;
           if pay_date < StrToDate('01.01.2000') then
             dept_name := 'unknown'
           else
             dept_name := disContractInvoiceInQuery.fieldbyname('dept_name').asstring;
//         comment := ;
           if accept = 'Y' then ContractDebitAccept := ContractDebitAccept + Amount;
           ContractDebitNoAccept := ContractDebitNoAccept + Amount;

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
           disContractInvoiceInQuery.Next;
         end;
         disContractInvoiceInQuery.Close;
         rowDebit := rowDebit + 2;

         while not disContractPlategiQueryDebitor.Eof do begin
           pay_date := disContractPlategiQueryDebitor.fieldbyname('doc_date').asdatetime;
//           cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
           doc_type := disContractPlategiQueryDebitor.fieldbyname('type_name').asstring;
           doc_no := disContractPlategiQueryDebitor.fieldbyname('pay_order').asstring;
//           short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
           amount := disContractPlategiQueryDebitor.fieldbyname('amount').asfloat;
           amount_usd := disContractPlategiQueryDebitor.fieldbyname('amount_usd').asfloat;
           contract_no := disContractPlategiQueryDebitor.fieldbyname('contract_no').asstring;
           accept := 'Y';
           comment := disContractPlategiQueryDebitor.fieldbyname('comment').asstring;
           if accept = 'Y' then contractDebitAccept := contractDebitAccept + Amount;
           contractDebitNoAccept := contractDebitNoAccept + Amount;

           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
//           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := cargo_date;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := doc_type;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := doc_no;
//           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := short_trade_mark;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := amount;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount_usd;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := contract_no;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := accept;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value := comment;

           rowDebit := rowDebit + 1;
           disContractPlategiQueryDebitor.Next;
         end;
         disContractPlategiQueryDebitor.Close;
         rowDebit := rowDebit + 2;

         while not disContractAnyQueryDebitor.Eof do begin
           pay_date := disContractAnyQueryDebitor.fieldbyname('pay_date').asdatetime;
//           cargo_date := disAnyQueryDebitor.fieldbyname('cargo_date').asdatetime;
           doc_type := disContractAnyQueryDebitor.fieldbyname('type_name').asstring;
//           doc_no := disAnyQueryDebitor.fieldbyname('act_no').asstring;
//           short_trade_mark := disAnyQueryDebitor.fieldbyname('short_trade_mark').asstring;
           amount := disContractAnyQueryDebitor.fieldbyname('amount').asfloat;
           amount_usd := disContractAnyQueryDebitor.fieldbyname('amount_usd').asfloat;
           contract_no := disContractAnyQueryDebitor.fieldbyname('contract_no').asstring;
           accept := 'Y';
           comment := disContractAnyQueryDebitor.fieldbyname('comment').asstring;
           if accept = 'Y' then contractDebitAccept := contractDebitAccept + Amount;
           contractDebitNoAccept := contractDebitNoAccept + Amount;

           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
//          vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := cargo_date;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := doc_type;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := doc_no;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := short_trade_mark;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := amount;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount_usd;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := contract_no;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := accept;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value := comment;

           rowDebit := rowDebit + 1;
           disContractAnyQueryDebitor.Next;
         end;
         disContractAnyQueryDebitor.Close;
         rowDebit := rowDebit + 2;

       { кредитовая статистика}
       { товарные отгрузки на предприятие }
         while not disContractInvoiceOutQuery.Eof do begin
           pay_date := disContractInvoiceOutQuery.fieldbyname('pay_date').asdatetime;
           invoice_date := disContractInvoiceOutQuery.fieldbyname('invoice_date').asdatetime;
           cargo_date := disContractInvoiceOutQuery.fieldbyname('cargo_date').asdatetime;
           doc_type := 'счет-фактура';
           doc_no := disContractInvoiceOutQuery.fieldbyname('invoice_no').asstring;
           short_trade_mark := disContractInvoiceOutQuery.fieldbyname('short_trade_mark').asstring;
           amount := disContractInvoiceOutQuery.fieldbyname('amount').asfloat;
           amount_usd := disContractInvoiceOutQuery.fieldbyname('amount_usd').asfloat;
           contract_no := disContractInvoiceOutQuery.fieldbyname('contract').asstring;
           accept := disContractInvoiceOutQuery.fieldbyname('is_in_oper').asstring;
           if pay_date < StrToDate('01.01.2000') then
             dept_name := 'unknown'
           else
             dept_name := disContractInvoiceOutQuery.fieldbyname('dept_name').asstring;
//         comment := ;
           if accept = 'Y' then contractCreditAccept := contractCreditAccept + Amount;
           contractCreditNoAccept := contractCreditNoAccept + Amount;

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
           disContractInvoiceOutQuery.Next;
         end;
         disContractInvoiceOutQuery.Close;
         rowCredit := rowCredit + 2;

         { вся кредитовая статистика кроме тов.отгрузок }
         while not disContractPlategiQueryCreditor.Eof do begin
           pay_date := disContractPlategiQueryCreditor.fieldbyname('doc_date').asdatetime;
//           cargo_date := disPlategiQueryCreditor.fieldbyname('cargo_date').asdatetime;
           doc_type := disContractPlategiQueryCreditor.fieldbyname('type_name').asstring;
           doc_no := disContractPlategiQueryCreditor.fieldbyname('pay_order').asstring;
//         short_trade_mark := disPlategiQueryCreditor.fieldbyname('short_trade_mark').asstring;
           amount := disContractPlategiQueryCreditor.fieldbyname('amount').asfloat;
           amount_usd := disContractPlategiQueryCreditor.fieldbyname('amount_usd').asfloat;
           contract_no := disContractPlategiQueryCreditor.fieldbyname('contract_no').asstring;
           accept := 'Y';
           comment := disContractPlategiQueryCreditor.fieldbyname('comment').asstring;
           if accept = 'Y' then contractCreditAccept := contractCreditAccept + Amount;
           contractCreditNoAccept := contractCreditNoAccept + Amount;

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
           disContractPlategiQueryCreditor.Next;
         end;
         disContractPlategiQueryCreditor.Close;
         rowCredit := rowCredit + 2;

         while not disContractAnyQueryCreditor.Eof do begin
           pay_date := disContractAnyQueryCreditor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
           doc_type := disContractAnyQueryCreditor.fieldbyname('type_name').asstring;
//          doc_no := disAnyQueryCreditor.fieldbyname('act_no').asstring;
//         short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
           amount := disContractAnyQueryCreditor.fieldbyname('amount').asfloat;
           amount_usd := disContractAnyQueryCreditor.fieldbyname('amount_usd').asfloat;
           contract_no := disContractAnyQueryCreditor.fieldbyname('contract_no').asstring;
           accept := 'Y';
           comment := disContractAnyQueryCreditor.fieldbyname('comment').asstring;
           if accept = 'Y' then contractCreditAccept := contractCreditAccept + Amount;
           contractCreditNoAccept := contractCreditNoAccept + Amount;

           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := pay_date;
//         vExcel.ActiveSheet.Cells[rowDebit,ColumnCredit + 1].Value := cargo_date;
           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := doc_type;
//         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := doc_no;
//         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := short_trade_mark;
           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := amount;
           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := amount_usd;
           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := contract_no;
           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 9].Value := accept;
           vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 11].Value := comment;

           rowCredit := rowCredit + 1;
           disContractAnyQueryCreditor.Next;
         end;
         disContractAnyQueryCreditor.Close;
         rowCredit := rowCredit + 2;

         { считаем сальдо  }
         contractSaldoNoAccept := contractSaldoBegin + contractDebitNoAccept - contractCreditNoAccept;
         contractSaldoAccept := contractSaldoBegin + contractDebitAccept - contractCreditAccept;

         all_contractes_saldo := all_contractes_saldo + contractSaldoAccept;

         if rowDebit > rowCredit then
           row := rowDebit
         else
           row := rowCredit;

         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'сальдо на конец периода:';
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := contractSaldoEnd;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'сальдо c акцептом:';
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := contractSaldoAccept;
         row := row + 1;
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 3].Value := 'сальдо без акцепта:';
         vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := contractSaldoNoAccept;
         row := row + 2;

         allEnterprContractQuery.Next;
       end; // конец    while not allEnterprContractQuery.Eof

       vExcel.ActiveSheet.Cells[row,3].Value :=
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------';

       RangeLeft := 'A' + IntToStr(row);
       RangeRight := 'Y' + IntToStr(row);
       vExcel.ActiveSheet.Range[OleVariant(RangeLeft),OleVariant(RangeRight)].Interior.ColorIndex := 6;

       // полное сальдо по предприятию для справки
       with allSaldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('saldo_date').asdate := EndDate;
         Open;
       end;
       all_saldo := allSaldoQuery.fieldbyname('allSaldo').asfloat;

       row := row + 3;
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 2].Value := 'СУММА САЛЬДО ПО ВСЕМ ДОГОВОРАМ:';
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := all_contractes_saldo;

       RangeLeft := 'C' + IntToStr(row);
       RangeRight := 'G' + IntToStr(row);

       if (round(all_contractes_saldo*100) <> round(all_saldo*100)) then
         vExcel.ActiveSheet.Range[OleVariant(RangeLeft),OleVariant(RangeRight)].Interior.ColorIndex := 3
       else
         vExcel.ActiveSheet.Range[OleVariant(RangeLeft),OleVariant(RangeRight)].Interior.ColorIndex := 35;

       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 2].Value := 'ПОЛНОЕ САЛЬДО С АКЦЕПТОМ ПО ПРЕДПРИЯТИЮ:';
       vExcel.ActiveSheet.Cells[row,ColumnDebit + 6].Value := all_saldo;

       // перемещаем указатель на следующее предприятие
       allContragentQuery.Next;
       Update;
     end;

   finally
     allSaldoQuery.Close;
     allContragentQuery.Close;
     allEnterprContractQuery.Close;
     contractSaldoQuery.Close;
     checkContractOperationQuery.Close;
     disContractPlategiQueryDebitor.Close;
     disContractPlategiQueryCreditor.Close;
     disContractAnyQueryDebitor.Close;
     disContractAnyQueryCreditor.Close;
     disContractInvoiceOutQuery.Close;
     disContractInvoiceInQuery.Close;
     vExcel := unAssigned;
     SetThreadLocale(Temp);
    end;
end;































//---------------------------------------------------------------------
// формирование отчета по всем операциям по ДИСу за указываемый период
//---------------------------------------------------------------------
procedure TStatisticReportForm.prepare_report_all_dis_operation;
Var
  old_lang : lcid;
  Excel : TExcel;
  PathToTemplate : string;
  ipID : integer;
const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
begin
  try
    old_lang := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sall_dis_operation;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

    // формируем отчет в Excel
    export_all_dis_operation(Excel);
    
  finally
    Excel.free;
    Excel := nil;
    SetThreadLocale(old_lang);
  end;
end;

// экспорт в Excel всех операций по ДИСу
procedure TStatisticReportForm.export_all_dis_operation(Excel : TExcel);
Var
  i : integer;
  row : integer;

  ent_id : integer;
  cur_ent_id : integer;
  prev_ent_id : integer;

  ReportHeader : string;
  cell : string;
  cellFrom : string;
  cellTo : string;
  s_row : string;
  info_row : array[1..15] of Variant;

  BeginDate : TDateTime;
  EndDate : TDateTime;

  is_coal_ent : string;
  last_coal_sender_date : TDate;
  ent_name : string;
  pay_date : TDate;
  cargo_date : TDate;
  invoice_date : TDate;
  doc_date : TDate;
  type_name : string;
  doc_no : string;
  short_trade_mark : string;
  amount : real;
  amount_usd : real;
  contract_no : string;
  accept : string;
  dept_name : string;
  comment : string;

begin
  BeginDate := StrToDate(all_opBeginMaskEdit.Text);
  EndDate := StrToDate(all_opEndMaskEdit.Text);

  try
    // просим все входящие операции
    with all_dis_in_operQuery do begin
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    all_dis_in_operQuery.open;

    // просим все исходящие операции
    with all_dis_out_operQuery do begin
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    all_dis_out_operQuery.Open;

    Excel.SelectWorkSheet('in_operation');
    ReportHeader := 'Все входящие операции по ДИСу за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate);
    row := 2;
    cell := 'A' + IntToStr(row);
    Excel.Cell[cell] := ReportHeader;

    row := 6;
    cur_ent_id := -1;
    prev_ent_id := -2;

    // ---- ---- ----- выполняем экспорт всех входящих оп-ций ----- ----- ----- //
    while not all_dis_in_operQuery.Eof do begin
      ent_id := all_dis_in_operQuery.fieldbyname('ent_id').asinteger;
      cur_ent_id := ent_id;
//      if cur_ent_id <> prev_ent_id then begin
//        // хотим понять явл-ся(лось) ли предприятие
//        // поставщиком угля
//        with is_coal_enterprQuery do begin
//          Close;
//          ParamByName('ent_id').asinteger := ent_id;
//        end;
//        is_coal_enterprQuery.Open;
//        is_coal_ent := is_coal_enterprQuery.fieldbyname('is_coal').asstring;
//        last_coal_sender_date := is_coal_enterprQuery.fieldbyname('last_sender_date').asdatetime;
//      end;
      ent_name := all_dis_in_operQuery.fieldbyname('enterpr_name').asstring;
      pay_date := all_dis_in_operQuery.fieldbyname('pay_date').asdatetime;
      cargo_date := all_dis_in_operQuery.fieldbyname('cargo_date').asdatetime;
      invoice_date := all_dis_in_operQuery.fieldbyname('invoice_date').asdatetime;
      type_name := all_dis_in_operQuery.fieldbyname('type_name').asstring;
      doc_no := all_dis_in_operQuery.fieldbyname('doc_no').asstring;
      short_trade_mark := all_dis_in_operQuery.fieldbyname('short_trade_mark').asstring;
      amount := all_dis_in_operQuery.fieldbyname('amount').asfloat;
      amount_usd := all_dis_in_operQuery.fieldbyname('amount_usd').asfloat;
      contract_no := all_dis_in_operQuery.fieldbyname('contract').asstring;
      accept := all_dis_in_operQuery.fieldbyname('accept').asstring;
      if pay_date < StrToDate('01.01.2000') then
        dept_name := 'unknown'
      else
        dept_name := all_dis_in_operQuery.fieldbyname('dept_name').asstring;
      comment := all_dis_in_operQuery.fieldbyname('comment').asstring;
      //-----------------------------------------------
      // начало блока экспорта в Excel
      info_row[1] := is_coal_ent;
      if is_coal_ent = 'Y' then
        info_row[2] := last_coal_sender_date
      else
        info_row[2] := '';
      info_row[3] := ent_name;
      info_row[4] := pay_date;

      if (cargo_date = 0) then
        info_row[5] := ''
      else
        info_row[5] := cargo_date;

      if (invoice_date = 0) then
        info_row[6] := ''
      else
        info_row[6] := invoice_date;

      info_row[7] := type_name;
      info_row[8] := doc_no;
      info_row[9] := short_trade_mark;
      info_row[10] := amount;
      info_row[11] := amount_usd;
      info_row[12] := contract_no;
      info_row[13] := accept;
      info_row[14] := dept_name;
      info_row[15] := comment;

      cellFrom := 'A' + IntToStr(row);
      cellTo := 'O' + IntToStr(row);
      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

      for i := 1 to 15 do info_row[i] := unAssigned;
      //-----------------------------------------------
      row := row + 1;
      prev_ent_id := cur_ent_id;
      all_dis_in_operQuery.Next;
    end;
    all_dis_in_operQuery.Close;

    Excel.SelectWorkSheet('out_operation');
    ReportHeader := 'Все исходящие операции по ДИСу за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate);
    row := 2;
    cell := 'A' + IntToStr(row);
    Excel.Cell[cell] := ReportHeader;

    row := 6;
    cur_ent_id := -1;
    prev_ent_id := -2;

    // ---- ---- ----- выполняем экспорт всех исходящих оп-ций ----- ----- ----- //
    while not all_dis_out_operQuery.Eof do begin
      ent_id := all_dis_out_operQuery.fieldbyname('ent_id').asinteger;
      cur_ent_id := ent_id;
//     if cur_ent_id <> prev_ent_id then begin
        // хотим понять явл-ся(лось) ли предприятие
        // поставщиком угля
//        with is_coal_enterprQuery do begin
//          Close;
//          ParamByName('ent_id').asinteger := ent_id;
//        end;
//        is_coal_enterprQuery.Open;
//        is_coal_ent := is_coal_enterprQuery.fieldbyname('is_coal').asstring;
//        last_coal_sender_date := is_coal_enterprQuery.fieldbyname('last_sender_date').asdatetime;
//      end;
      ent_name := all_dis_out_operQuery.fieldbyname('enterpr_name').asstring;
      pay_date := all_dis_out_operQuery.fieldbyname('pay_date').asdatetime;
      cargo_date := all_dis_out_operQuery.fieldbyname('cargo_date').asdatetime;
      invoice_date := all_dis_out_operQuery.fieldbyname('invoice_date').asdatetime;
      type_name := all_dis_out_operQuery.fieldbyname('type_name').asstring;
      doc_no := all_dis_out_operQuery.fieldbyname('doc_no').asstring;
      short_trade_mark := all_dis_out_operQuery.fieldbyname('short_trade_mark').asstring;
      amount := all_dis_out_operQuery.fieldbyname('amount').asfloat;
      amount_usd := all_dis_out_operQuery.fieldbyname('amount_usd').asfloat;
      contract_no := all_dis_out_operQuery.fieldbyname('contract').asstring;
      accept := all_dis_out_operQuery.fieldbyname('accept').asstring;
      if pay_date < StrToDate('01.01.2000') then
        dept_name := 'unknown'
      else
        dept_name := all_dis_out_operQuery.fieldbyname('dept_name').asstring;
      comment := all_dis_out_operQuery.fieldbyname('comment').asstring;
      //-----------------------------------------------
      // начало блока экспорта в Excel
      info_row[1] := is_coal_ent;
      if is_coal_ent = 'Y' then
        info_row[2] := last_coal_sender_date
      else
        info_row[2] := '';
      info_row[3] := ent_name;
      info_row[4] := pay_date;
      if (cargo_date = 0) then
        info_row[5] := ''
      else
        info_row[5] := cargo_date;

      if (invoice_date = 0) then
        info_row[6] := ''
      else
        info_row[6] := invoice_date;

      info_row[7] := type_name;
      info_row[8] := doc_no;
      info_row[9] := short_trade_mark;
      info_row[10] := amount;
      info_row[11] := amount_usd;
      info_row[12] := contract_no;
      info_row[13] := accept;
      info_row[14] := dept_name;
      info_row[15] := comment;

      cellFrom := 'A' + IntToStr(row);
      cellTo := 'O' + IntToStr(row);
      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

      for i := 1 to 15 do info_row[i] := unAssigned;
      //-----------------------------------------------
      row := row + 1;
      prev_ent_id := cur_ent_id;
      all_dis_out_operQuery.Next;
    end;
    all_dis_out_operQuery.Close;

   finally
     is_coal_enterprQuery.Close;
     all_dis_in_operQuery.Close;
     all_dis_out_operQuery.Close;
   end;
end;
//---------------------------------------------------------------------
// конец формирование отчета по всем операциям по ДИСу за указываемый период
//---------------------------------------------------------------------






















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
     contract : string;  // используется при plategiCheckBox.Checked = true;
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

     row := 4;
     if plategiCheckBox.Checked then
       vExcel.ActiveSheet.Cells[row, 11].Value := 'Договор';

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
       contract := allPlategiQuery.fieldbyname('contract_no').asstring;
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
       if plategiCheckBox.Checked then
         vExcel.ActiveSheet.Cells[row,Column + 10].Value := contract;
       vExcel.ActiveSheet.Cells[row,Column + 13].Value := comment;

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

//---------------------------------------------------------------------
// формирование статистики по предприятию за указываемый период
// форма - для составления баланса
//---------------------------------------------------------------------
procedure TStatisticReportForm.ExportStatisticNew;
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
     trade_mark : string;
     amount : real;
     amount_usd : real;
     contract_no : string;
     accept : string;
     dept_name : string;
     comment : string;

     invoice_id : integer;
     nds : real;
     cargo_sender_name : string;
     cargo_receiver_name : string;
     qnty : real;
     full_price : real;
     price_without_nds : real;
     sum_without_nds : real;
     full_sum : real;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     ColumnName = 11;
     ColumnDebit = ColumnName + 1;
     ColumnCredit = 1;
  begin
     temp := GetThreadLocale;
     SetThreadLocale(English_Locale);
     pname := @s;

     BeginDate := StrToDate(BeginNewMaskEdit.Text);
     EndDate := StrToDate(EndNewMaskEdit.Text);

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
{     with disInvoiceInQuery do begin
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
}
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
     PathToTemplate := PathToProgram + '\Template\' + sStatisticNewTemplate;
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

     with allInvQuery do begin
       Close;
       SQL.Clear;
       SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
       SQL.Add('where sender_id = :id and payer_id = 0');
       SQL.Add('order by sender_name, pay_date');
       Prepare;
     end;

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

       with allInvQuery do begin
         Close;
         ParamByName('id').asfloat := ent_id;
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
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := 'Сальдо на начало периода:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 6].Value := allSaldoBegin;
       row := row + 2;
       rowDebit := row;
       rowCredit := row;

       { формирование дебитовой статистики}

       { отгрузка с предприятия }
       while not allInvQuery.eof do begin
         invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
         pay_date := allInvQuery.fieldbyname('pay_date').asdatetime;
         doc_no := allInvQuery.fieldbyname('invoice_no').asstring;
         amount := allInvQuery.fieldbyname('amount').asfloat;
         cargo_sender_name := allInvQuery.fieldbyname('cargo_sender').asstring;
         cargo_receiver_name := allInvQuery.fieldbyname('cargo_receiver').asstring;
         cargo_date := allInvQuery.fieldbyname('cargo_date').asdatetime;
         invoice_date := allInvQuery.fieldbyname('invoice_date').asdatetime;
         if pay_date < StrToDate('01.01.2000') then
            dept_name := 'unknown'
         else
            dept_name := allInvQuery.fieldbyname('dept_name').asstring;
         contract_no := allInvQuery.fieldbyname('contract').asstring;
         amount_usd := allInvQuery.fieldbyname('amount_usd').asfloat;
         accept := allInvQuery.fieldbyname('is_in_oper').asstring;

         if accept = 'Y' then allDebitAccept := allDebitAccept + Amount;
         allDebitNoAccept := allDebitNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := cargo_sender_name;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := cargo_receiver_name;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 10].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 11].Value := invoice_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 12].Value := dept_name;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 13].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 14].Value := amount_usd;

         // detail
         with InvoiceItemsQuery do begin
           Close;
           ParamByName('invoice_id').asinteger := invoice_id;
           Open;
         end;

         with ExtraInvoiceItemsQuery do begin
           Close;
           ParamByName('invoice_id').asinteger := invoice_id;
           Open;
         end;

         // invoice items
         while not InvoiceItemsQuery.eof do begin
           trade_mark := InvoiceItemsQuery.fieldbyname('trade_mark').asstring;
           qnty := InvoiceItemsQuery.fieldbyname('qnty').asfloat;
           sum_without_nds := InvoiceItemsQuery.fieldbyname('summ_without_nds').asfloat;
           full_sum := InvoiceItemsQuery.fieldbyname('full_summ').asfloat;
           full_price := InvoiceItemsQuery.fieldbyname('full_price').asfloat;

           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := trade_mark;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 3].Value := qnty;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := sum_without_nds;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := full_sum;
           if qnty <> 0 then
             vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 6].Value := Real(full_sum)/Real(qnty);
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := cargo_sender_name;
           vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := cargo_receiver_name;

           rowDebit := rowDebit + 1;
           InvoiceItemsQuery.Next;
         end;

       // extra invoice items
       while not ExtraInvoiceItemsQuery.eof do begin
         trade_mark := ExtraInvoiceItemsQuery.fieldbyname('extra_item_name').asstring;
         price_without_nds := ExtraInvoiceItemsQuery.fieldbyname('price_without_nds').asfloat;
         full_price := ExtraInvoiceItemsQuery.fieldbyname('full_price').asfloat;
         //
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := trade_mark;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 4].Value := price_without_nds;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 5].Value := full_price;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 8].Value := cargo_sender_name;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 9].Value := cargo_receiver_name;

         rowDebit := rowDebit + 1;
         ExtraInvoiceItemsQuery.Next;
       end;

         rowDebit := rowDebit + 1;
         allInvQuery.Next;
       end;
       allInvQuery.Close;
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
         comment := disPlategiQueryDebitor.fieldbyname('comment').asstring;
         allDebitNoAccept := allDebitNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 13].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 14].Value := amount_usd;

         rowDebit := rowDebit + 1;
         disPlategiQueryDebitor.Next;
       end;
       disPlategiQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       while not disAnyQueryDebitor.Eof do begin
         pay_date := disAnyQueryDebitor.fieldbyname('pay_date').asdatetime;
         doc_no := disAnyQueryDebitor.fieldbyname('act_no').asstring;
         doc_type := disAnyQueryDebitor.fieldbyname('type_name').asstring;
         amount := disAnyQueryDebitor.fieldbyname('amount').asfloat;
         amount_usd := disAnyQueryDebitor.fieldbyname('amount_usd').asfloat;
         contract_no := disAnyQueryDebitor.fieldbyname('contract_no').asstring;
         comment := disAnyQueryDebitor.fieldbyname('comment').asstring;
         allDebitNoAccept := allDebitNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 0].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 1].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 2].Value := doc_type;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 7].Value := amount;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 13].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowDebit,ColumnDebit + 14].Value := amount_usd;

         rowDebit := rowDebit + 1;
         disAnyQueryDebitor.Next;
       end;
       disAnyQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       { кредитовая статистика}
       { товарные отгрузки на предприятие }
       while not disInvoiceOutQuery.Eof do begin
         cargo_date := disInvoiceOutQuery.fieldbyname('cargo_date').asdatetime;
         invoice_date := disInvoiceOutQuery.fieldbyname('invoice_date').asdatetime;
         pay_date := disInvoiceOutQuery.fieldbyname('pay_date').asdatetime;
         if pay_date < StrToDate('01.01.2000') then
            dept_name := 'unknown'
         else
            dept_name := disInvoiceOutQuery.fieldbyname('dept_name').asstring;
         contract_no := disInvoiceOutQuery.fieldbyname('contract').asstring;
         amount_usd := disInvoiceOutQuery.fieldbyname('amount_usd').asfloat;
         doc_no := disInvoiceOutQuery.fieldbyname('invoice_no').asstring;
         amount := disInvoiceOutQuery.fieldbyname('amount').asfloat;
         nds := disInvoiceOutQuery.fieldbyname('nds').asfloat;
         trade_mark := disInvoiceOutQuery.fieldbyname('short_trade_mark').asstring;
         accept := disInvoiceOutQuery.fieldbyname('is_in_oper').asstring;
//         comment := ;
         if accept = 'Y' then allCreditAccept := allCreditAccept + Amount;
         allCreditNoAccept := allCreditNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 0].Value := cargo_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 1].Value := invoice_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 2].Value := dept_name;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 3].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 5].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := amount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := amount-nds;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 9].Value := trade_mark;

         rowCredit := rowCredit + 1;
         disInvoiceOutQuery.Next;
       end;
       disInvoiceOutQuery.Close;
       rowCredit := rowCredit + 2;

       { вся кредитовая статистика кроме тов.отгрузок }
       while not disPlategiQueryCreditor.Eof do begin
         pay_date := disPlategiQueryCreditor.fieldbyname('doc_date').asdatetime;
         doc_type := disPlategiQueryCreditor.fieldbyname('type_name').asstring;
         doc_no := disPlategiQueryCreditor.fieldbyname('pay_order').asstring;
         amount := disPlategiQueryCreditor.fieldbyname('amount').asfloat;
         amount_usd := disPlategiQueryCreditor.fieldbyname('amount_usd').asfloat;
         contract_no := disPlategiQueryCreditor.fieldbyname('contract_no').asstring;
//         comment := disPlategiQueryCreditor.fieldbyname('comment').asstring;
         allCreditNoAccept := allCreditNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 5].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := amount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 10].Value := doc_type;

         rowCredit := rowCredit + 1;
         disPlategiQueryCreditor.Next;
       end;
       disPlategiQueryCreditor.Close;
       rowCredit := rowCredit + 2;

       while not disAnyQueryCreditor.Eof do begin
         contract_no := disAnyQueryCreditor.fieldbyname('contract_no').asstring;
         amount_usd := disAnyQueryCreditor.fieldbyname('amount_usd').asfloat;
         pay_date := disAnyQueryCreditor.fieldbyname('pay_date').asdatetime;
         doc_no := disAnyQueryCreditor.fieldbyname('act_no').asstring;
         amount := disAnyQueryCreditor.fieldbyname('amount').asfloat;
         doc_type := disAnyQueryCreditor.fieldbyname('type_name').asstring;
         allCreditNoAccept := allCreditNoAccept + Amount;

         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 4].Value := contract_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 5].Value := amount_usd;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 6].Value := pay_date;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 7].Value := doc_no;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 8].Value := amount;
         vExcel.ActiveSheet.Cells[rowCredit,ColumnCredit + 10].Value := doc_type;

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
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := 'Полное сальдо на конец периода:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 6].Value := allSaldoEnd;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := 'Полное сальдо c акцептом:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 6].Value := allSaldoAccept;
       row := row + 1;
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 3].Value := 'Полное сальдо без акцепта:';
       vExcel.ActiveSheet.Cells[row,ColumnCredit + 6].Value := allSaldoNoAccept;
       row := row + 1;

       // перемещаем указатель на следующее предприятие
       allContragentQuery.Next;
       Update;
     end;

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
// формирование сводного отчета по полученным углям с разбивкой
// по предприятиям-поставщикам и грузополучателям
//---------------------------------------------------------------------

function CoalIDToCoalName(supply_id : integer):string;
Var temp : string;
begin
  temp := '';
  case supply_id of
    iG     : temp := 'Г';       // марка Г
    iK     : temp := 'К';       // марка К
    iDG    : temp := 'ДГ';      // марка ДГ
    iOC    : temp := 'ОС';      // марка ОС
    iGG    : temp := 'Ж';       // марка Ж
    iDGKOM : temp := 'ДГКОМ';   // марка ДГКОМ
    iT     : temp := 'Т';       // марка Т
    iAK    : temp := 'АК';      // марка АК
    iAKO   : temp := 'АКО';     // марка АКО
    iAO    : temp := 'АО';      // марка АО
    iAC    : temp := 'АС';      // марка АС
    iGr    : temp := 'Гр';      // марка Гр коксующийся
    iGGr   : temp := 'Жр';      // марка Жр
    iKr    : temp := 'Кр';      // марка Кр
    iOCr   : temp := 'ОСр';     // марка ОСр
    iTr    : temp := 'Тр';      // марка Тр коксующийся
  end;
  CoalIDToCoalName := temp;
end;

// процедура выполняет расчет данных для отчета
// и формирует определенную структуру данных
// которая затем исп-ся для построения отчета
procedure TStatisticReportForm.CreateCoalSenderReport(Var Coal : CoalSender);
Var
  i,i1 : integer;
  //
  prev_cargo_id, cur_cargo_id : real;
  prev_prod_id, cur_prod_id : integer;
  prev_supply_id, cur_supply_id : integer;
  prev_ref, cur_ref : string;
  //
  cargo_receiver : string;
  coal_name : string;
  qnty : real;
  summ_without_nds : real;
  full_summ : real;
  nds : real;

begin
  // инициализация массивов для сводного угольного баланса
  Coal.count := 0;
  for i := 1 to iMaxItems do begin
    Coal.Group[i].count := 0;
    Coal.Group[i].all_pripl_skidki := 0;
    Coal.Group[i].all_sum_free_vat := 0;
    Coal.Group[i].all_nds := 0;
    for i1 := 1 to iMaxItems do begin
      Coal.Group[i].Coal[i1].coal_name_id := 0;
      Coal.Group[i].Coal[i1].coal_name := '';
      Coal.Group[i].Coal[i1].qnty := 0;
      Coal.Group[i].Coal[i1].pure_sum_free_vat := 0;
      Coal.Group[i].Coal[i1].add_pripl_skidki := 0;
      Coal.Group[i].Coal[i1].nds := 0;
      Coal.Group[i].Coal[i1].cargo_receiver := '';
      Coal.Group[i].Coal[i1].receiver := '';
    end;
  end;
  //
  prev_cargo_id := -1;
  cur_cargo_id := -1;
  prev_prod_id := -1;
  cur_prod_id := -1;
  prev_supply_id := -1;
  cur_supply_id := -1;
//  prev_ref := '';
  cur_ref := '';
  //
  while not coalSenderInvQuery.Eof do begin
    cur_ref := coalSenderInvQuery.fieldbyname('is_in_ref').asstring;
    cur_cargo_id := coalSenderInvQuery.fieldbyname('CARGO_RECEIVER_ID').asinteger;
    cur_prod_id := coalSenderInvQuery.fieldbyname('PROD_ID').asinteger;
    cur_supply_id := coalSenderInvQuery.fieldbyname('SUPPLY_ID').asinteger;

    cargo_receiver := coalSenderInvQuery.fieldbyname('CARGO_RECEIVER').asstring;
    coal_name := CoalIDToCoalName(cur_supply_id);
    qnty := coalSenderInvQuery.fieldbyname('qnty').asfloat;
    summ_without_nds := coalSenderInvQuery.fieldbyname('summ_without_nds').asfloat;
    full_summ := coalSenderInvQuery.fieldbyname('full_summ').asfloat;
    nds := full_summ - summ_without_nds;

    // обрабатываем давальческий счет
    if cur_ref = 'N' then begin
    end;

    // обрабатываем счет к-рый будет перевыставляться
    if cur_ref = 'Y' then begin
    end;

    prev_ref := cur_ref;
    prev_cargo_id := cur_cargo_id;
    prev_prod_id := cur_prod_id;
    prev_supply_id := cur_supply_id;

    //  переходим на следующую запись
    coalSenderInvQuery.Next;
  end;
end;

procedure TStatisticReportForm.ExportCoalReport(Sender: TObject);
  Var
     temp: lcid;
     vExcel : Variant;

     BeginDate : TDateTime;
     EndDate : TDateTime;
     PathToTemplate : string;
     i : integer;
//     ReportHeader : string;
     row : integer;
     rowDetail : integer;
     cellFrom : string;
     cellTo : string;

     { контрольные переменные }
     allCoalQnty : real;
     allCoalAmountFreeVAT : real;
     allCoalAmountVAT : real;
     countSender : integer ;

     // invoice master
     sender_id : real;
     sender_name : string;
     payer_name : string;
     amount : real;
     nds : real;
     cargo_receiver_name : string;
     qnty : real;

 const
    English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
    Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    PathToTemplate := PathToProgram + '\Template\' + sCoalReportTemplate;
    try
      vExcel.AddWorkBook(PathToTemplate);
      vExcel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try
     row := 2;
//     cell := 'A' + IntToStr(row);
//     Excel.Cell[cell] := ReportHeader;

{     BeginDate := StrToDate(BeginCoalReportMaskEdit.Text);
     EndDate := StrToDate(EndCoalReportMaskEdit.Text);

     with allCoalSenderQuery do begin
       Close;
       SQL.Clear;
       SQL.Add('select distinct sender_id, enterpr_name');
       SQL.Add('from balans_report_input_coal_all(:begin_date, :end_date)');
       SQL.Add('where prod_id = 10010 or prod_id = 10011');
       SQL.Add('order by enterpr_name');
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
     end;

     coalSenderInvQuery.Prepare;

     { инициализируем  контрольные переменные }
{     countSender := 0;
     row := 6;

     { просим в базе поставщиков углей }
{     allCoalSenderQuery.Open;

  // ---- ---- ----- начало цикла по поставщикам угля ----- ----- ----- //
     while not allCoalSenderQuery.Eof do begin
       countSender := countSender + 1;
       rowDetail := row; // запоминаем значение строки для связанных счетов

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       // master
       // экспортирем наименование предприятия
       cellFrom := 'A' + IntToStr(row);
       cellTo := 'A' + IntToStr(row);
       sender_id := allCoalSenderQuery.fieldbyname('sender_id').asfloat;
       sender_name := allCoalSenderQuery.fieldbyname('enterpr_name').asstring;
       info_row[1] := sender_name;
       vExcel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       // формируем отчет по углю для предприятия с кодом sender_id
       with coalSenderInvQuery do begin
         ParamByName('sender_id').asfloat := sender_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
       end;
       coalSenderInvQuery.Open;
       // формируем из запроса массив с информацией об угле
       CreateCoalSenderReport();
       coalSenderInvQuery.Close;





       payer_name := allInvQuery.fieldbyname('payer_name').asstring;
       amount := allInvQuery.fieldbyname('amount').asfloat;
       nds := allInvQuery.fieldbyname('nds').asfloat;
       cargo_receiver_name := allInvQuery.fieldbyname('cargo_receiver').asstring;

//       s_row := IntToStr(row);
       info_row[1] := countInvoices;
       info_row[2] := sender_name;
       info_row[3] := payer_name;
       info_row[4] := pay_date;
       info_row[5] := invoice_date;
       info_row[6] := cargo_date;
       info_row[7] := invoice_no;
       info_row[15] := amount;
       info_row[16] := nds;
       info_row[17] := amount_usd;
       info_row[18] := cargo_sender_name;
       info_row[19] := cargo_receiver_name;
       info_row[20] := contract_no;
       info_row[21] := accept;
       info_row[22] := dept_name;
       info_row[23] := act_no;
       if Double(act_date) <> 0 then
         info_row[24] := act_date
       else
         info_row[24] := '';
       info_row[25] := inv_type;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'AF' + IntToStr(row);
//       Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
       //       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       // detail
       with InvoiceItemsQuery do begin
         Close;
         ParamByName('invoice_id').asinteger := invoice_id;
         Open;
       end;









       ItemsFlag := false;
       // invoice items
       while not InvoiceItemsQuery.eof do begin
         trade_mark := InvoiceItemsQuery.fieldbyname('trade_mark').asstring;
         dimention := InvoiceItemsQuery.fieldbyname('dimention').asstring;
         qnty := InvoiceItemsQuery.fieldbyname('qnty').asfloat;
         price_without_nds := InvoiceItemsQuery.fieldbyname('price_without_nds').asfloat;
         full_price := InvoiceItemsQuery.fieldbyname('full_price').asfloat;
         sum_without_nds := InvoiceItemsQuery.fieldbyname('summ_without_nds').asfloat;
         full_sum := InvoiceItemsQuery.fieldbyname('full_summ').asfloat;
         //
         info_row[8] := trade_mark;
         info_row[9] := dimention;
         info_row[10] := qnty;
         info_row[11] := price_without_nds;
         info_row[12] := full_price;
         info_row[13] := sum_without_nds;
         info_row[14] := full_sum;
         // добавляем грузополучателей для каждой строки invoice_items
         info_row[2] := sender_name;
         info_row[3] := payer_name;
         info_row[4] := pay_date;
         info_row[5] := invoice_date;
         info_row[6] := cargo_date;
         info_row[7] := invoice_no;
         info_row[18] := cargo_sender_name;
         info_row[19] := cargo_receiver_name;
         info_row[20] := contract_no;
         info_row[21] := accept;
         info_row[22] := dept_name;
         info_row[23] := act_no;
         if Double(act_date) <> 0 then
           info_row[24] := act_date
         else
           info_row[24] := '';
         info_row[25] := inv_type;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'AF' + IntToStr(row);
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         ItemsFlag := true;
         vExcel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 32 do info_row[i] := unAssigned;

         row := row + 1;
         InvoiceItemsQuery.Next;
       end;

//       for i := 1 to 32 do info_row[i] := '';

       ExtraItemsFlag := false;
       // extra invoice items
       while not ExtraInvoiceItemsQuery.eof do begin
         trade_mark := ExtraInvoiceItemsQuery.fieldbyname('extra_item_name').asstring;
         price_without_nds := ExtraInvoiceItemsQuery.fieldbyname('price_without_nds').asfloat;
         full_price := ExtraInvoiceItemsQuery.fieldbyname('full_price').asfloat;
         //
         info_row[8] := trade_mark;
         info_row[13] := price_without_nds;
         info_row[14] := full_price;
         // добавляем грузополучателей для каждой строки invoice_items
         info_row[2] := sender_name;
         info_row[3] := payer_name;
         info_row[4] := pay_date;
         info_row[5] := invoice_date;
         info_row[6] := cargo_date;
         info_row[7] := invoice_no;
         info_row[18] := cargo_sender_name;
         info_row[19] := cargo_receiver_name;
         info_row[20] := contract_no;
         info_row[21] := accept;
         info_row[22] := dept_name;
         info_row[23] := act_no;
         if Double(act_date) <> 0 then
           info_row[24] := act_date
         else
           info_row[24] := '';
         info_row[25] := inv_type;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'AF' + IntToStr(row);
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         ExtraItemsFlag := true;
         vExcel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 32 do info_row[i] := unAssigned;

         row := row + 1;
         ExtraInvoiceItemsQuery.Next;
       end;

       if (ItemsFlag = false) and (ExtraItemsFlag = false) then
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         vExcel.Range[cellFrom,cellTo] := VarArrayOf(info_row);


       end;

       row := row + 1;
       allCoalSenderQuery.Next;
     end;
}
   finally
     vExcel.free;
//     allCoalSenderQuery.Close;
//     ExtraInvoiceItemsQuery.Close;
//     InvoiceItemsQuery.Close;
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
  if StatisticPageControl.ActivePage.Name = sStatForBalansPageName then
    ExportStatisticNew;
  if StatisticPageControl.ActivePage.Name = sCoalBalansPageName then
    ExportCoalStatistic;
  if StatisticPageControl.ActivePage.Name = sall_opPageName then
    prepare_report_all_dis_operation;
//  if StatisticPageControl.ActivePage.Name = sCoalReportPageName then
//    ExportCoalReport;
  // 
  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TStatisticReportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TStatisticReportForm.AllCoalCheckBoxClick(Sender: TObject);
begin
  if (AllCoalCheckBox.Checked) then begin
    CoalOnlyCheckBox.Visible := true;
    CoalOnlyCheckBox.Enabled := true;
    end
  else begin
    CoalOnlyCheckBox.Visible := false;
    CoalOnlyCheckBox.Enabled := false;
    end
end;

procedure TStatisticReportForm.StatisticCheckBoxClick(Sender: TObject);
begin
  if StatisticCheckBox.Checked then
    contractCheckBox.Enabled := false
  else
    contractCheckBox.Enabled := true;
end;

procedure TStatisticReportForm.contractCheckBoxClick(Sender: TObject);
begin
  if contractCheckBox.Checked then begin
    StatisticCheckBox.Enabled := false;
    detailInvCheckBox.Enabled := false;
    allContractCheckBox.Enabled := false;
  end
  else begin
//    StatisticCheckBox.Enabled := true;
    detailInvCheckBox.Enabled := true;
    allContractCheckBox.Enabled := true;
  end;
end;

procedure TStatisticReportForm.allContractCheckBoxClick(Sender: TObject);
begin
  if allContractCheckBox.Checked then begin
    StatisticCheckBox.Enabled := false;
    detailInvCheckBox.Enabled := false;
    contractCheckBox.Enabled := false;
  end
  else begin
//    StatisticCheckBox.Enabled := true;
    detailInvCheckBox.Enabled := true;
    contractCheckBox.Enabled := true;
  end;
end;

procedure TStatisticReportForm.FormHide(Sender: TObject);
Var
  BeginDate, EndDate : TDateTime;
begin
  BeginDate := StrToDate(StatBeginMaskEdit.Text);
  EndDate := StrToDate(StatEndMaskEdit.Text);
  parentConfig.SharedDll.WriteDate(BeginDate,EndDate);
end;

end.
