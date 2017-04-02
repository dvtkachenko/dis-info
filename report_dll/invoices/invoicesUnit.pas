unit invoicesUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, shared_type, excel_type;

const
  iEnterprPage = 0;
  iDeptPage = 1;
  iAllInvPage = 2;
  iCoalCoxInvPage = 3;
  sforEnterprInvPage = 'forEnterprTabSheet';
  sforDeptInvPage = 'forDeptTabSheet';
  sAllInvPage = 'allInvTabSheet';
  sCoalCoxInvPage = 'coalcoxInvTabSheet';
  sproductionPage = 'productionTabSheet';
  sInvoicesTemplate = 'invoices_detail.xlt';

type
  TInvExportForm = class(TForm)
    InvPageControl: TPageControl;
    allInvQuery: TQuery;
    forDeptTabSheet: TTabSheet;
    InvoiceItemsQuery: TQuery;
    ExtraInvoiceItemsQuery: TQuery;
    InvBeginMaskEdit: TMaskEdit;
    InvEndMaskEdit: TMaskEdit;
    InOutRadioGroup: TRadioGroup;
    forEnterprTabSheet: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    chainInvQuery: TQuery;
    ruleGroupBox: TGroupBox;
    chainCheckBox: TCheckBox;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    allInvTabSheet: TTabSheet;
    Label5: TLabel;
    Label7: TLabel;
    coalcoxInvTabSheet: TTabSheet;
    Label8: TLabel;
    Label9: TLabel;
    SkidkiPriplQuery: TQuery;
    ccRadioGroup: TRadioGroup;
    SkidkiPriplCheckBox: TCheckBox;
    zdtarifCheckBox: TCheckBox;
    zdtarifQuery: TQuery;
    coal_cox_weightQuery: TQuery;
    is_coal_invQuery: TQuery;
    our_tarifQuery: TQuery;
    is_cox_invQuery: TQuery;
    addPanel: TPanel;
    conditionRadioGroup: TRadioGroup;
    Bevel1: TBevel;
    addtoListBitBtn: TBitBtn;
    productionTabSheet: TTabSheet;
    Label10: TLabel;
    Label11: TLabel;
    conditionListBox: TListBox;
    dopinfoCheckBox: TCheckBox;
    nalogovayaQuery: TQuery;
    get_user_nameQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure NotDetailInvoicesReport(Sender: TObject);
    procedure InitExcel;
    procedure DeInitExcel;
    procedure DetailInvoicesReport(Excel : TExcel);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure coalcoxInvTabSheetShow(Sender: TObject);
    procedure coalcoxInvTabSheetHide(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure zdtarifCheckBoxClick(Sender: TObject);
    procedure SkidkiPriplCheckBoxClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure conditionRadioGroupClick(Sender: TObject);
    procedure productionTabSheetShow(Sender: TObject);
    procedure productionTabSheetHide(Sender: TObject);
    procedure addtoListBitBtnClick(Sender: TObject);
    procedure prepareQuery_for_prod_inv;
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    // содержит список id в строковом виде для соответствующих
    // наименований продукции из conditionListBox
    // данные в списке содержатся в отсортированном виде ,
    // поэтому установить соответствие между именем продукции
    // в conditionListBox и номерами в conditionListBox
    conditionValueList: TStringList;
  public
    { Public declarations }
    parentConfig : p_config;
    ReportHeader : string;
    BeginDate : TDateTime;
    EndDate : TDateTime;
    PathToProgram : string;
    Excel : TExcel;
    old_lang: lcid;
    InterprocessCall : boolean;
    // параметры , инициализируемые вызываемой DLL 
    ipID : integer;
    ipExcel : TExcel;
  end;

implementation

{$R *.DFM}

function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';
function GetProduction(const mode:integer; Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetProduction';

{сервисные процедуры}

{-------------------}

procedure TInvExportForm.FormShow(Sender: TObject);
Var
  BeginDate, EndDate : TDateTime;
begin
  parentConfig.SharedDll.ReadDate(BeginDate,EndDate);
  InvBeginMaskEdit.Text := DateToStr(BeginDate);
  InvEndMaskEdit.Text := DateToStr(EndDate);
  if InvPageControl.ActivePage.Name <> sCoalCoxInvPage then begin
    SkidkiPriplCheckBox.Checked := false;
    SkidkiPriplCheckBox.Enabled := false;
    SkidkiPriplCheckBox.Visible := false;
  end;
  //
  if InvPageControl.ActivePage.Name = sproductionPage then begin
    conditionRadioGroup.ItemIndex := -1;
    conditionListBox.Items.Clear;
    conditionValueList.Clear;
    conditionValueList.Duplicates := dupError;
    conditionValueList.Sorted := true;
    addtoListBitBtn.Enabled := false;
    Height := 490;
    sbReportToExcel.Enabled := false;
    // поле Tag используется для хранения предыдущего значения
    // ItemIndex
    // инициализируем заведомо не совпадающим значением 
    conditionRadioGroup.Tag := -2;
  end
  else begin
    Height := 306;
    sbReportToExcel.Enabled := true;
  end;
end;

//---------------------------------------------------------------------
// инициализация Excel
//---------------------------------------------------------------------
procedure TInvExportForm.InitExcel;
Var
  PathToTemplate : string;
const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
begin
  if Excel = nil then begin
    old_lang := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sInvoicesTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;  
  end;
end;

//---------------------------------------------------------------------
// деинициализация Excel
//---------------------------------------------------------------------
procedure TInvExportForm.DeInitExcel;
begin
  if Excel <> nil then begin
    Excel.free;
    Excel := nil;
    SetThreadLocale(old_lang);
  end;
end;

//---------------------------------------------------------------------
// выбор счетов-фактур закрепленных за отделами (не детализированный отчет)
//---------------------------------------------------------------------
procedure TInvExportForm.NotDetailInvoicesReport(Sender: TObject);
{  Var
     temp: lcid;
     vExcel : Variant;
     BeginDate : TDateTime;
     EndDate : TDateTime;
     PathToTemplate : string;
     ReportHeader : string;
     dept_id : real;
     dept_name : string;
     row : integer;
     Column : integer;

     { контрольные переменные }
{     allInvoiceAmountAccepted : real;
     allInvoiceAmountUsdAccepted : real;
     allInvoiceAmountNotAccepted : real;
     allInvoiceAmountUsdNotAccepted : real;
     countInvoices : integer ;

     sender_name : string;
     payer_name : string;
     pay_date : TDate;
     invoice_date : TDate;
     cargo_date : TDate;
     invoice_no : string;
     short_trade_mark : string;
     amount : real;
     amount_usd : real;
     nds : real;
     contract_no : string;
     accept : string;
     cargo_sender_name : string;
     cargo_receiver_name : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
}  begin
{     temp := GetThreadLocale;
     SetThreadLocale(English_Locale);

     BeginDate := StrToDate(InvBeginMaskEdit.Text);
     EndDate := StrToDate(InvEndMaskEdit.Text);

     { конструирование запросов }

{     if InOutRadioGroup.ItemIndex = 0 then
       begin
         ReportHeader := 'Входящие ';
         with allInvQuery do begin
           Close;
           SQL.Clear;
           SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
           SQL.Add('where dept_id = :dept and payer_id = 0');
           SQL.Add('order by sender_name, pay_date');
         end
       end;

     if InOutRadioGroup.ItemIndex = 1 then
       begin
         ReportHeader := 'Исходящие ';
         with allInvQuery do begin
           Close;
           SQL.Clear;
           SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
           SQL.Add('where dept_id = :dept and sender_id = 0');
           SQL.Add('order by payer_name, pay_date');
         end
       end;
{
     if FindEnterpriseForm.ShowModal = mrOk then
       begin
         dept_id := FindEnterpriseForm.FindEnterpriseQuery.fieldbyname('object_id').asfloat;
       end
     else
         raise Exception.Create('Отдел не выбран');
     end;
//     ParamByName('begin_date').asdate := BeginDate;
//     ParamByName('end_date').asdate := EndDate;
}
{     try
     	vExcel := GetActiveOleObject('Excel.Application');
     except
       try
         vExcel := CreateOleObject('Excel.Application');
       except
         raise Exception.Create('Невозможно загрузить Excel');
       end;
     end;
     vExcel.Visible := true;

   try
     PathToTemplate := PathToProgram + '\Template\invoices_detail.xls';
     vExcel.Application.Workbooks.Open(PathToTemplate);
     ReportHeader := ReportHeader + 'счета-фактуры за период с ' +
                      datetostr(BeginDate) + ' по ' + datetostr(EndDate) +
                      ' ' + '(' + dept_name  + ')';
     row := 2;
     vExcel.ActiveSheet.Cells[row, 1].Value := ReportHeader;

     { инициализируем  контрольные переменные }
{     allInvoiceAmountAccepted := 0;
     allInvoiceAmountUsdAccepted := 0;
     allInvoiceAmountNotAccepted := 0;
     allInvoiceAmountUsdNotAccepted := 0;
     countInvoices := 0;
     row := 7;

     { просим в базе необходимые счета }
{     allInvQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
{     while not allInvQuery.Eof do begin
       countInvoices := countInvoices + 1;

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       sender_name := allInvQuery.fieldbyname('sender_name').asstring;
       payer_name := allInvQuery.fieldbyname('payer_name').asstring;
       pay_date := allInvQuery.fieldbyname('pay_date').asdatetime;
       invoice_date := allInvQuery.fieldbyname('invoice_date').asdatetime;
       cargo_date := allInvQuery.fieldbyname('cargo_date').asdatetime;
       invoice_no := allInvQuery.fieldbyname('invoice_no').asstring;
       short_trade_mark := allInvQuery.fieldbyname('short_trade_mark').asstring;
       amount := allInvQuery.fieldbyname('amount').asfloat;
       amount_usd := allInvQuery.fieldbyname('amount_usd').asfloat;
       nds := allInvQuery.fieldbyname('nds').asfloat;
       contract_no := allInvQuery.fieldbyname('contract').asstring;
       accept := allInvQuery.fieldbyname('is_in_oper').asstring;
       cargo_sender_name := allInvQuery.fieldbyname('cargo_sender').asstring;
       cargo_receiver_name := allInvQuery.fieldbyname('cargo_receiver').asstring;





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
}
{       row := row + 1;
       allInvQuery.Next;
     end;

   finally
     allInvQuery.Close;
     vExcel := unAssigned;
     SetThreadLocale(Temp);
    end;}
end;

//---------------------------------------------------------------------
// выбор счетов-фактур закрепленных за отделами (детализированный отчет)
//---------------------------------------------------------------------
procedure TInvExportForm.DetailInvoicesReport(Excel : TExcel);
  Var
     temp: lcid;
     cell : string;
     cellFrom : string;
     cellTo : string;
     s_row : string;
     info_row : array[1..34] of Variant;
     i : integer;
     ItemsFlag, ExtraItemsFlag : boolean;
//     ReportHeader : string;
     dept_id : integer;
     row : integer;
     rowDetail : integer;
     rowChain : integer;
     rowDopInfo : integer;

     { контрольные переменные }
//     allInvoiceAmountAccepted : real;
//     allInvoiceAmountUsdAccepted : real;
//     allInvoiceAmountNotAccepted : real;
//     allInvoiceAmountUsdNotAccepted : real;
     allInvoiceAmount : real;
     allInvoiceItemsAmount : real;
     allInvoiceItemsAmountFreeVat : real;
     allInvoiceNDS : real;
     countInvoices : integer ;
     countInvSubItem : integer ;
     countChainInv : integer ;
     countNakladnaya : integer ;

     is_coal_inv : string;
     is_cox_inv : string;

     // invoice master
     invoice_id : integer;
     sender_name : string;
     payer_name : string;
     pay_date : TDate;
     invoice_date : TDate;
     cargo_date : TDate;
     invoice_no : string;
     amount : real;
     amount_usd : real;
     skidki_pripl : real;  // необходима для угольного отчета
     zdtarif : real;  // необходима для угольного отчета
     our_tarif : string;  // признаки ж/д тарифа
     nds : real;
     coal_cox_weight : real;  // необходима для угольного отчета
     coal_cox_dry_weight : real;  // необходима для угольного отчета
     is_correct : string;  // необходима для угольного отчета
     cargo_sender_name : string;
     cargo_receiver_name : string;
     contract_no : string;
     accept : string;
     reference : string;
     prim : string;
     act_no : string;
     act_date : TDate;
     inv_type : string;
     dept_name : string;
     user_name : string;
     // invoice detail
     trade_mark : string;
     dimention : string;
     qnty : real;
     price_without_nds : real;
     full_price : real;
     sum_without_nds : real;
     full_sum : real;
     // chain invoices
     c_invoice_no : string;
     c_pay_date : TDate;
     c_invoice_date : TDate;
     c_contract_no : string;
     c_sender_name : string;
     c_payer_name : string;
     c_amount : real;      // добавлено 26.10.2007 для одного отчета
     // налоговые накладные
     nakladnaya_no : string;
     nakladnaya_date : TDate;
     nakladnaya_sum : real;

  begin

   try
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     // если выбраны "приплаты/скидки" , то меняем соответствующим образом
     // заголовок столбца в отчете
     if (InvPageControl.ActivePage.TabIndex = iCoalCoxInvPage) and
        (SkidkiPriplCheckBox.Checked) then begin
       cell := 'Q' + IntToStr(4);
       Excel.Cell[cell] := 'Приплаты/Скидки без НДС';
     end;

     // если выбран "показать ж/д тариф" , то меняем соответствующим образом
     // заголовок столбца в отчете
     if zdtarifCheckBox.Checked then begin
       cell := 'Q' + IntToStr(4);
       Excel.Cell[cell] := 'Ж/д тариф по квитанции б/НДС';
     end;

     { инициализируем  контрольные переменные }
//     allInvoiceAmountUsdAccepted := 0;
//     allInvoiceAmountNotAccepted := 0;
//     allInvoiceAmountUsdNotAccepted := 0;
     countInvoices := 0;
     row := 6;

     { просим в базе необходимые счета }
     allInvQuery.Open;
     InvoiceItemsQuery.Prepare;
     ExtraInvoiceItemsQuery.Prepare;
     zdtarifQuery.Prepare;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allInvQuery.Eof do begin
       countInvoices := countInvoices + 1;
       // переменная для подсчета элементов в счете
       countInvSubItem := 0;
       rowChain := row; // запоминаем значение строки для связанных счетов
       rowDopInfo := row; // запоминаем значение строки для дополнительной информации
       // переменные для проверки совпадения суммы продукции по счету с суммой счета
       allInvoiceAmount := 0;
       allInvoiceItemsAmount := 0;
       allInvoiceItemsAmountFreeVat := 0;
       allInvoiceNDS := 0;

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       // master
       invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;

       sender_name := allInvQuery.fieldbyname('sender_name').asstring;
       payer_name := allInvQuery.fieldbyname('payer_name').asstring;
       pay_date := allInvQuery.fieldbyname('pay_date').asdatetime;
       invoice_date := allInvQuery.fieldbyname('invoice_date').asdatetime;
       cargo_date := allInvQuery.fieldbyname('cargo_date').asdatetime;
       invoice_no := allInvQuery.fieldbyname('invoice_no').asstring;
       invoice_no := ' ' + invoice_no;
       amount := allInvQuery.fieldbyname('amount').asfloat;
       nds := allInvQuery.fieldbyname('nds').asfloat;
       amount_usd := allInvQuery.fieldbyname('amount_usd').asfloat;
       cargo_sender_name := allInvQuery.fieldbyname('cargo_sender').asstring;
       cargo_receiver_name := allInvQuery.fieldbyname('cargo_receiver').asstring;
       contract_no := allInvQuery.fieldbyname('contract').asstring;
       accept := allInvQuery.fieldbyname('is_in_oper').asstring;
       dept_name := allInvQuery.fieldbyname('dept_name').asstring;
       prim := allInvQuery.fieldbyname('prim').asstring;
       act_no := allInvQuery.fieldbyname('act_no').asstring;
       act_date := allInvQuery.fieldbyname('act_date').asdatetime;
       inv_type := allInvQuery.fieldbyname('programm_type').asstring;
       if inv_type = 'D' then inv_type := 'дав';
       if inv_type = 'U' then inv_type := 'обычн';
       if inv_type = 'E' then inv_type := 'экспорт';
       reference := allInvQuery.fieldbyname('is_in_ref').asstring;

       allInvoiceAmount := allInvoiceAmount + amount;
       allInvoiceNDS := allInvoiceNDS + nds;

       //  формируем столбец приплаты/скидки
       if (InvPageControl.ActivePage.TabIndex = iCoalCoxInvPage) and
          (SkidkiPriplCheckBox.Checked) then begin
         with SkidkiPriplQuery do begin
           Close;
           ParamByName('invoice_id').asinteger := invoice_id;
           Open;
         end;
         skidki_pripl := SkidkiPriplQuery.fieldbyname('skidki_pripl').asfloat;
         //  подменяем значение amount_usd на приплаты скидки
         //  теперь у нас приплаты/скидки будут выводиться в
         //  отдельном столбце для каждого счета
         amount_usd := skidki_pripl;
       end;

       // формируем столбец ж/д тариф
       if zdtarifCheckBox.Checked then begin
         with zdtarifQuery do begin
           Close;
           ParamByName('invoice_id').asinteger := invoice_id;
           Open;
         end;
         zdtarif := zdtarifQuery.fieldbyname('sum_tarif').asfloat;
         // подменяем значение amount_usd на ж/д тариф
         amount_usd := zdtarif;

         with our_tarifQuery do begin
           Close;
           ParamByName('invoice_id').asinteger := invoice_id;
           Open;
         end;
         our_tarif := our_tarifQuery.fieldbyname('our_tarif').asstring;
         cellFrom := 'Q' + IntToStr(row);
         cellTo := 'Q' + IntToStr(row);
         // красим в зеленый цвет ячейку, если ж/д тариф сторонний
         if (our_tarif = 'N') then Excel.FillRangeColor(cellFrom,cellTo,43);
         // красим в желтый цвет ячейку, если в базу не по каждой квитанции
         // занесен признак стороннего или своего тарифа
         if (our_tarif = 'B') then Excel.FillRangeColor(cellFrom,cellTo,6);
         // красим в голубой цвет ячейку, если в счете есть тариф и наш
         // и сторонний
         if (our_tarif = 'I') then Excel.FillRangeColor(cellFrom,cellTo,34);
       end;

       // до 01.10.2000 г качество по сертификатам еще не разносилось
       // поэтому смысл во всех проверках в счетах до 01.10.2000 отпадает
       if (pay_date >= StrToDate('01.10.2000')) then begin
         // угольный ли счет
         with is_coal_invQuery do begin
           Close;
           ParamByName('inv_id').asinteger := invoice_id;
           Open;
         end;
         is_coal_inv := is_coal_invQuery.fieldbyname('is_coal').asstring;

         // коксовый ли счет
         with is_cox_invQuery do begin
           Close;
           ParamByName('inv_id').asinteger := invoice_id;
           Open;
         end;
         is_cox_inv := is_cox_invQuery.fieldbyname('is_cox').asstring;

         if ((is_coal_inv = 'Y') or (is_cox_inv = 'Y')) then begin
           with coal_cox_weightQuery do begin
             Close;
             ParamByName('inv_id').asinteger := invoice_id;
             Open;
           end;
           coal_cox_weight := coal_cox_weightQuery.fieldbyname('weight').asfloat;
           coal_cox_dry_weight := coal_cox_weightQuery.fieldbyname('dry_weight').asfloat;
           is_correct := coal_cox_weightQuery.fieldbyname('is_correct').asstring;
           info_row[18] := coal_cox_weight;
           info_row[19] := coal_cox_dry_weight;
           // красим в красный цвет ячейку, если сухой вес возможно не верен
           cellFrom := 'S' + IntToStr(row);
           cellTo := 'S' + IntToStr(row);
           if (is_correct = 'N') then Excel.FillRangeColor(cellFrom,cellTo,46);
         end;
       end;

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
       info_row[20] := cargo_sender_name;
       info_row[21] := cargo_receiver_name;
       info_row[22] := contract_no;
       info_row[23] := accept;
       info_row[24] := dept_name;
       info_row[25] := act_no;
       if Double(act_date) <> 0 then
         info_row[26] := act_date
       else
         info_row[26] := '';
       info_row[27] := inv_type;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'AH' + IntToStr(row);
//       Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
       //       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

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
         allInvoiceItemsAmount := allInvoiceItemsAmount + full_sum;
         allInvoiceItemsAmountFreeVat := allInvoiceItemsAmountFreeVat + sum_without_nds;

         countInvSubItem := countInvSubItem + 1;

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
         info_row[20] := cargo_sender_name;
         info_row[21] := cargo_receiver_name;
         info_row[22] := contract_no;
         info_row[23] := accept;
         info_row[24] := dept_name;
         info_row[25] := act_no;
         if Double(act_date) <> 0 then
           info_row[26] := act_date
         else
           info_row[26] := '';
         info_row[27] := inv_type;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'AH' + IntToStr(row);
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         ItemsFlag := true;
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         // красим в желтый цвет неакцептованные счета
         if (accept = 'N') then Excel.FillRangeColor(cellFrom,cellTo,6);

         for i := 1 to 34 do info_row[i] := unAssigned;

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
         allInvoiceItemsAmount := allInvoiceItemsAmount + full_price;
         allInvoiceItemsAmountFreeVat := allInvoiceItemsAmountFreeVat + price_without_nds;

         countInvSubItem := countInvSubItem + 1;

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
         info_row[20] := cargo_sender_name;
         info_row[21] := cargo_receiver_name;
         info_row[22] := contract_no;
         info_row[23] := accept;
         info_row[24] := dept_name;
         info_row[25] := act_no;
         if Double(act_date) <> 0 then
           info_row[26] := act_date
         else
           info_row[26] := '';
         info_row[27] := inv_type;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'AH' + IntToStr(row);
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         ExtraItemsFlag := true;
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         // красим в желтый цвет неакцептованные счета
         if (accept = 'N') then Excel.FillRangeColor(cellFrom,cellTo,6);

         for i := 1 to 34 do info_row[i] := unAssigned;

         row := row + 1;
         ExtraInvoiceItemsQuery.Next;
       end;

       if (ItemsFlag = false) and (ExtraItemsFlag = false) then begin
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         countInvSubItem := countInvSubItem + 1;
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         // красим в желтый цвет неакцептованные счета
         if (accept = 'N') then Excel.FillRangeColor(cellFrom,cellTo,6);
         row := row + 1;
       end;

       cellFrom := 'O' + IntToStr(row-1);
       cellTo := 'Q' + IntToStr(row-1);
       if ((Round(allInvoiceItemsAmount*100) <> Round(allInvoiceAmount*100)) or
           (Round((allInvoiceItemsAmountFreeVat+allInvoiceNDS)*100)
            <> Round(allInvoiceAmount*100))
          ) then begin
         // красим ячейки в красный если указанные суммы не свпадают
         Excel.FillRangeColor(cellFrom,cellTo,46);
         cell := 'Q' + IntToStr(row-1);
         Excel.Cell[cell] := 'не совпадает сумма счета!!!';
       end;


       // экспортируем связанные счета
       if (chainCheckBox.Checked) and (reference = 'Y')then begin

         for i := 1 to 34 do info_row[i] := unAssigned;

         with chainInvQuery do begin
           Close;
           ParamByName('inv_id').asinteger := invoice_id;
           Open;
         end;

         // обнуляем счетчик привязанных счетов
         countChainInv := 0;
         while not chainInvQuery.eof do begin
           c_invoice_no := chainInvQuery.fieldbyname('invoice_no').asstring;
           c_invoice_no := ' ' + c_invoice_no;
           c_pay_date := chainInvQuery.fieldbyname('pay_date').asdatetime;
           c_invoice_date := chainInvQuery.fieldbyname('invoice_date').asdatetime;
           c_contract_no := chainInvQuery.fieldbyname('contract_no').asstring;
           c_sender_name := chainInvQuery.fieldbyname('sender_name').asstring;
           c_payer_name := chainInvQuery.fieldbyname('payer_name').asstring;
           c_amount := chainInvQuery.fieldbyname('amount').asfloat;

           // увеличиваем счетчик привязанных счетов на 1
           countChainInv := countChainInv + 1;
           info_row[1] := c_invoice_no;
           info_row[2] := c_pay_date;
           info_row[3] := c_invoice_date;
           info_row[4] := c_contract_no;
           info_row[5] := c_sender_name;
           info_row[6] := c_payer_name;
//           info_row[7] := c_amount;  // добавлено 26.10.2007 для одного отчета

           cellFrom := 'AC' + IntToStr(rowChain);
           cellTo := 'AH' + IntToStr(rowChain);
//           cellTo := 'AL' + IntToStr(rowChain); // добавлено 26.10.2007 для одного отчета
//           Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
           Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

           rowChain := rowChain + 1;
           chainInvQuery.Next;
         end;
         // для корректного использования фильтра по связанным счетам
         // выравниваем число строк в связанных счетах
         if countChainInv = 1 then begin
           for i := 1 to countInvSubItem - 1 do begin
             cellFrom := 'AC' + IntToStr(rowChain);
             cellTo := 'AH' + IntToStr(rowChain);
//           cellTo := 'AL' + IntToStr(rowChain); // добавлено 26.10.2007 для одного отчета
             Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

             rowChain := rowChain + 1;
           end;
         end;

         if row < rowChain then row := rowChain;

       end;  //  конец экспорта связанных счетов

       // экспортируем дополнительную информацию о счете
       if (dopinfoCheckBox.Checked) then begin

         for i := 1 to 34 do info_row[i] := unAssigned;

         with nalogovayaQuery do begin
           Close;
           ParamByName('inv_id').asinteger := invoice_id;
           Open;
         end;

         with get_user_nameQuery do begin
           Close;
           ParamByName('inv_id').asinteger := invoice_id;
           Open;
         end;
         user_name := get_user_nameQuery.fieldbyname('user_name').asstring;

         // обнуляем счетчик налоговых накладных
         countNakladnaya := 0;
         while not nalogovayaQuery.eof do begin
           nakladnaya_no := nalogovayaQuery.fieldbyname('nakladnaya_no').asstring;
           nakladnaya_no := '' + nakladnaya_no;
           nakladnaya_date := nalogovayaQuery.fieldbyname('nakladnaya_date').asdatetime;
           nakladnaya_sum := nalogovayaQuery.fieldbyname('summa').asfloat;

           countNakladnaya := countNakladnaya + 1;

           info_row[1] := nakladnaya_no;
           info_row[2] := nakladnaya_date;
           info_row[3] := nakladnaya_sum;
           info_row[4] := unAssigned;
           info_row[5] := invoice_id;
           info_row[6] := user_name;
           info_row[7] := prim; // примечание к счету

           cellFrom := 'AJ' + IntToStr(rowDopInfo);
           cellTo := 'AP' + IntToStr(rowDopInfo);
           Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

           rowDopInfo := rowDopInfo + 1;

           nalogovayaQuery.Next;
         end;

         if countNakladnaya = 0 then begin
           info_row[4] := unAssigned;
           info_row[5] := invoice_id;
           info_row[6] := user_name;
         end;

         // для корректного использования фильтра по налоговым накладным
         // выравниваем число строк в налоговых накладных
         if countNakladnaya <= 1 then begin
           for i := 1 to countInvSubItem - countNakladnaya do begin
             cellFrom := 'AJ' + IntToStr(rowDopInfo);
             cellTo := 'AP' + IntToStr(rowDopInfo);
             // делаем чтобы сумма не задваивалась,
             // а показывалась только 1 раз
             info_row[3] := ' ';

             Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

             rowDopInfo := rowDopInfo + 1;
           end;
         end;

         if row < rowDopInfo then row := rowDopInfo;

       end;
       row := row + 1;
       allInvQuery.Next;
     end;

   finally
     allInvQuery.Close;
     ExtraInvoiceItemsQuery.Close;
     InvoiceItemsQuery.Close;
     zdtarifQuery.Close;
     our_tarifQuery.Close;
     nalogovayaQuery.Close;
     get_user_nameQuery.Close;
   end;
end;

//------------------------------------------------------------
// процедура формирования запроса для вытяжки
// счетов-фактур по продукции или по ее типам
//------------------------------------------------------------
procedure TInvExportForm.prepareQuery_for_prod_inv;
var
  i : integer;
  str_id : string;
begin
  // формирование запроса для вытяжки по типу продукции
  if conditionRadioGroup.ItemIndex = 0 then begin
    with allInvQuery do begin
      Close;
      SQL.Clear;
      SQL.Add('select distinct b.* from balans_report_all_invoices(:begin_date, :end_date) b,');
      SQL.Add('invoice_items it, supply s');
      // если счет входящий, то
      if InOutRadioGroup.ItemIndex = 0 then begin
        SQL.Add('where b.payer_id = 0');
        ReportHeader := 'Входящие ';
      end;

      // если счет исходящий, то
      if InOutRadioGroup.ItemIndex = 1 then begin
        SQL.Add('where b.sender_id = 0');
        ReportHeader := 'Исходящие ';
      end;

      SQL.Add('and it.invoice_id = b.invoice_id');
      SQL.Add('and it.supply_id = s.supply_id');
      SQL.Add('and (');
      // вытаскиваем 1-ый элемент из списка значений
      str_id := conditionValueList.Strings[0];
      SQL.Add('s.prod_id = ' + str_id);
      // вытаскиваем следующие элементы
      for i := 1 to conditionValueList.Count - 1 do begin
        str_id := conditionValueList.Strings[i];
        SQL.Add('or s.prod_id = ' + str_id);
      end;
      SQL.Add(')');
      // если счет входящий, то
      if InOutRadioGroup.ItemIndex = 0 then
        SQL.Add('order by b.sender_name, b.pay_date');
      // если счет исходящий, то
      if InOutRadioGroup.ItemIndex = 1 then
        SQL.Add('order by b.payer_name, b.pay_date');
      Prepare;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end
  end;

  // формирование запроса для вытяжки по наименованию продукции
  if conditionRadioGroup.ItemIndex = 1 then begin
    with allInvQuery do begin
      Close;
      SQL.Clear;
      SQL.Add('select distinct b.* from balans_report_all_invoices(:begin_date, :end_date) b,');
      SQL.Add('invoice_items it');
      // если счет входящий, то
      if InOutRadioGroup.ItemIndex = 0 then begin
        SQL.Add('where b.payer_id = 0');
        ReportHeader := 'Входящие ';
      end;

      // если счет исходящий, то
      if InOutRadioGroup.ItemIndex = 1 then begin
        SQL.Add('where b.sender_id = 0');
        ReportHeader := 'Исходящие ';
      end;

      SQL.Add('and it.invoice_id = b.invoice_id');
      SQL.Add('and (');
      // вытаскиваем 1-ый элемент из списка значений
      str_id := conditionValueList.Strings[0];
      SQL.Add('it.supply_id = ' + str_id);
      // вытаскиваем следующие элементы
      for i := 1 to conditionValueList.Count - 1 do begin
        str_id := conditionValueList.Strings[i];
        SQL.Add('or it.supply_id = ' + str_id);
      end;
      SQL.Add(')');
      // если счет входящий, то
      if InOutRadioGroup.ItemIndex = 0 then
        SQL.Add('order by b.sender_name, b.pay_date');
      // если счет исходящий, то
      if InOutRadioGroup.ItemIndex = 1 then
        SQL.Add('order by b.payer_name, b.pay_date');
      Prepare;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end
  end;

  // формирование запроса для вытяжки по наименованию
  // доп.расходов
  if conditionRadioGroup.ItemIndex = 2 then begin
    with allInvQuery do begin
      Close;
      SQL.Clear;
      SQL.Add('select distinct b.* from balans_report_all_invoices(:begin_date, :end_date) b,');
      SQL.Add('invoice_extra_items it');
      // если счет входящий, то
      if InOutRadioGroup.ItemIndex = 0 then begin
        SQL.Add('where b.payer_id = 0');
        ReportHeader := 'Входящие ';
      end;

      // если счет исходящий, то
      if InOutRadioGroup.ItemIndex = 1 then begin
        SQL.Add('where b.sender_id = 0');
        ReportHeader := 'Исходящие ';
      end;

      SQL.Add('and it.invoice_id = b.invoice_id');
      SQL.Add('and (');
      // вытаскиваем 1-ый элемент из списка значений
      str_id := conditionValueList.Strings[0];
      SQL.Add('it.extra_id = ' + str_id);
      // вытаскиваем следующие элементы
      for i := 1 to conditionValueList.Count - 1 do begin
        str_id := conditionValueList.Strings[i];
        SQL.Add('or it.extra_id = ' + str_id);
      end;
      SQL.Add(')');
      Prepare;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end
  end;

end;
//------------------------------------------------------------

procedure TInvExportForm.sbReportToExcelClick(Sender: TObject);
Var
  id : integer;
  name : string;
  s : array[0..maxPChar] of Char;
  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
  pname := @s;
  BeginDate := StrToDate(InvBeginMaskEdit.Text);
  EndDate := StrToDate(InvEndMaskEdit.Text);

  if (InvPageControl.ActivePage.Name = sforDeptInvPage) then
       begin
         if InOutRadioGroup.ItemIndex = 0 then begin
           ReportHeader := 'Входящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
             SQL.Add('where dept_id = :dept and payer_id = 0');
             SQL.Add('order by sender_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end
         end;

         if InOutRadioGroup.ItemIndex = 1 then begin
           ReportHeader := 'Исходящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
             SQL.Add('where dept_id = :dept and sender_id = 0');
             SQL.Add('order by payer_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end
         end;

         if GetDepatment(id,pname) = mrOk then begin
           name := string(pname);
           allInvQuery.ParamByName('dept').asinteger := id;
         end
         else
          raise Exception.Create('Отдел не выбран');

       end; // конец sforDeptInvPage

  if (InvPageControl.ActivePage.Name = sforEnterprInvPage) then
       begin
         if InOutRadioGroup.ItemIndex = 0 then begin
           ReportHeader := 'Входящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
             SQL.Add('where sender_id = :id and payer_id = 0');
             SQL.Add('order by sender_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end
         end;

         if InOutRadioGroup.ItemIndex = 1 then begin
           ReportHeader := 'Исходящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
             SQL.Add('where payer_id = :id and sender_id = 0');
             SQL.Add('order by payer_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end
         end;

         if not InterprocessCall then begin
           if GetEnterprise(id,pname) = mrOk then begin
             name := string(pname);
             allInvQuery.ParamByName('id').asinteger := id;
           end
           else
            raise Exception.Create('Предприятие не выбрано');
         end
         else begin
           // инициализируем параметром , переданным из другой DLL
           allInvQuery.ParamByName('id').asinteger := ipID;
         end;


       end; // конец sforEnterprInvPage

  if (InvPageControl.ActivePage.Name = sAllInvPage) then
       begin
         if InOutRadioGroup.ItemIndex = 0 then begin
           ReportHeader := 'Входящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
             SQL.Add('where payer_id = 0');
             SQL.Add('order by sender_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end
         end;

         if InOutRadioGroup.ItemIndex = 1 then begin
           ReportHeader := 'Исходящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
             SQL.Add('where sender_id = 0');
             SQL.Add('order by payer_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end
         end;
         name := 'все счета по ДИСу';

       end; // конец sAllInvPage

  if (InvPageControl.ActivePage.Name = sCoalCoxInvPage) then
       begin
         if InOutRadioGroup.ItemIndex = 0 then begin
           ReportHeader := 'Входящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select distinct i.* from');
             SQL.Add('balans_report_all_invoices(:begin_date, :end_date) i,');
             SQL.Add('invoice_items it, supply s');
             SQL.Add('where i.payer_id = 0 and');
             SQL.Add('i.invoice_id = it.invoice_id and');
             SQL.Add('it.supply_id = s.supply_id and');
             if ccRadioGroup.ItemIndex = 0 then begin
               SQL.Add('(s.prod_id = 10010 or s.prod_id = 10011)');
             end;
             if ccRadioGroup.ItemIndex = 1 then begin
               SQL.Add('(s.prod_id = 13)');
             end;
             SQL.Add('order by sender_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end;
           //
           if ccRadioGroup.ItemIndex = 0 then
             name := 'уголь';
           if ccRadioGroup.ItemIndex = 1 then
             name := 'кокс';
         end;

         if InOutRadioGroup.ItemIndex = 1 then begin
           ReportHeader := 'Исходящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add('select distinct i.* from');
             SQL.Add('balans_report_all_invoices(:begin_date, :end_date) i,');
             SQL.Add('invoice_items it, supply s');
             SQL.Add('where i.sender_id = 0 and');
             SQL.Add('i.invoice_id = it.invoice_id and');
             SQL.Add('it.supply_id = s.supply_id and');
             if ccRadioGroup.ItemIndex = 0 then begin
               SQL.Add('(s.prod_id = 10010 or s.prod_id = 10011)');
             end;
             if ccRadioGroup.ItemIndex = 1 then begin
               SQL.Add('(s.prod_id = 13)');
             end;
             SQL.Add('order by payer_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end;
           //
           if ccRadioGroup.ItemIndex = 0 then
             name := 'уголь';
           if ccRadioGroup.ItemIndex = 1 then
             name := 'кокс';
         end;

       end; // конец sCoalCoxInvPage

  if (InvPageControl.ActivePage.Name = sproductionPage) then
       begin
         prepareQuery_for_prod_inv;
       end; // конец sproductionPage

  // формируем запрос для связываемых счетов
  if chainCheckBox.Checked then begin
    with chainInvQuery do begin
      Close;
      SQL.Clear;
{      SQL.Add('select * from balans_invoices_list(0) il,');
      SQL.Add('invoice_references ir');
      if InOutRadioGroup.ItemIndex = 0 then begin
        SQL.Add('where il.invoice_id = ir.out_id');
        SQL.Add('and ir.in_id = :inv_id');
      end;
      if InOutRadioGroup.ItemIndex = 1 then begin
        SQL.Add('where il.invoice_id = ir.in_id');
        SQL.Add('and ir.out_id = :inv_id');
      end;
      SQL.Add('order by pay_date');
}
      SQL.Add('select i.invoice_no, i.rate_date pay_date, i.invoice_date,');
      SQL.Add('s.enterprise_name sender_name, p.enterprise_name payer_name,');
      SQL.Add('i.contract_no,');
      SQL.Add('i.amount ');
      SQL.Add('from invoices i,');
      SQL.Add('invoice_references ir,');
      SQL.Add('enterpr s,');
      SQL.Add('enterpr p');
      if InOutRadioGroup.ItemIndex = 0 then begin
        SQL.Add('where i.invoice_id = ir.out_id');
        SQL.Add('and ir.in_id = :inv_id');
      end;
      if InOutRadioGroup.ItemIndex = 1 then begin
        SQL.Add('where i.invoice_id = ir.in_id');
        SQL.Add('and ir.out_id = :inv_id');
      end;
      SQL.Add('and i.sender_id = s.enterpr_id');
      SQL.Add('and i.payer_id = p.enterpr_id');
      SQL.Add('order by i.rate_date');
      Prepare;
    end  // end of WITH
  end;   // end of IF
  ReportHeader := ReportHeader + 'счета-фактуры за период с ' +
                  InvBeginMaskEdit.Text + ' по ' + InvEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
  if not InterprocessCall then begin
    InitExcel;
    DetailInvoicesReport(Excel);
    DeInitExcel;
  end
  else begin
    DetailInvoicesReport(ipExcel);
  end;

  if not InterprocessCall then begin
    Application.BringToFront;
    MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
  end;  
end;

procedure TInvExportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TInvExportForm.coalcoxInvTabSheetShow(Sender: TObject);
begin
  SkidkiPriplCheckBox.Enabled := true;
  SkidkiPriplCheckBox.Visible := true;
end;

procedure TInvExportForm.coalcoxInvTabSheetHide(Sender: TObject);
begin
  SkidkiPriplCheckBox.Checked := false;
  SkidkiPriplCheckBox.Enabled := false;
  SkidkiPriplCheckBox.Visible := false;
end;

procedure TInvExportForm.FormHide(Sender: TObject);
Var
  BeginDate, EndDate : TDateTime;
begin
  BeginDate := StrToDate(InvBeginMaskEdit.Text);
  EndDate := StrToDate(InvEndMaskEdit.Text);
  parentConfig.SharedDll.WriteDate(BeginDate,EndDate);
end;

procedure TInvExportForm.zdtarifCheckBoxClick(Sender: TObject);
begin
  if SkidkiPriplCheckBox.Checked then begin
    zdtarifCheckBox.Checked := false;
  end;
end;

procedure TInvExportForm.SkidkiPriplCheckBoxClick(Sender: TObject);
begin
  if zdtarifCheckBox.Checked then begin
    SkidkiPriplCheckBox.Checked := false;
  end;
end;

procedure TInvExportForm.FormCreate(Sender: TObject);
begin
  Excel := nil;
  old_lang := 0;
  // создаем список для хранения значений
  // используемых при выборе счетов-фактур по
  // типу продукции, ее названию или по доп.расходам
  conditionValueList := TStringList.Create;
end;

procedure TInvExportForm.FormDestroy(Sender: TObject);
begin
  // удаляем список для хранения значений
  // используемых при выборе счетов-фактур по
  // типу продукции, ее названию или по доп.расходам 
  conditionValueList.Free;
end;

// делает недоступной кнопку "Добавить в список" если не выбрано
// ни одно условие поиска
procedure TInvExportForm.conditionRadioGroupClick(Sender: TObject);
begin
  // сравниваем предыдущее значение ItemIndex с текущим
  // используем для хранения предыдущего значение
  // св-во Tag
  // если prev <> cur , то очищаем список
  if conditionRadioGroup.Tag <> conditionRadioGroup.ItemIndex then begin
    conditionRadioGroup.Tag := conditionRadioGroup.ItemIndex;
    conditionListBox.Items.Clear;
    conditionValueList.Clear;
    if conditionListBox.Items.Count = 0 then
      sbReportToExcel.Enabled := false;
  end;

  if conditionRadioGroup.ItemIndex = -1 then
    addtoListBitBtn.Enabled := false
  else
    addtoListBitBtn.Enabled := true;
end;

// показываем дополнительные условия если выбрана
// закладка "по продукции"
procedure TInvExportForm.productionTabSheetShow(Sender: TObject);
begin
  // поле Tag используется для хранения предыдущего значения
  // ItemIndex
  // инициализируем заведомо не совпадающим значением
  conditionRadioGroup.Tag := -2;

  conditionRadioGroup.ItemIndex := -1;
  conditionListBox.Items.Clear;
  conditionValueList.Clear;
  conditionValueList.Duplicates := dupError;
  conditionValueList.Sorted := true;
  addtoListBitBtn.Enabled := false;
  Height := 490;
  if conditionListBox.Items.Count = 0 then
    sbReportToExcel.Enabled := false;
end;

procedure TInvExportForm.productionTabSheetHide(Sender: TObject);
begin
  // поле Tag используется для хранения предыдущего значения
  // ItemIndex
  // инициализируем заведомо не совпадающим значением
  conditionRadioGroup.Tag := -2;

  Height := 306;
  sbReportToExcel.Enabled := true;
end;

// обработчик нажатия кнопки "Добавить в список"
procedure TInvExportForm.addtoListBitBtnClick(Sender: TObject);
Var
  mode : integer;
  id : integer;
  name : string;
  s : array[0..maxPChar] of Char;
  pname : PChar;
//  productionItem : TListItem;
begin
  pname := @s;
//  productionItem := nil;

  try
    // производим вызов сервисной процедуры поиска
    // по типу продукции
    if conditionRadioGroup.ItemIndex = 0 then begin
      mode := iprod_type_mode;
      if GetProduction(mode,id,pname) = mrOk then
        name := string(pname)
      else
        raise Exception.Create('Не выбран тип продукции');
    end;

    // производим вызов сервисной процедуры поиска
    // по названию продукции
    if conditionRadioGroup.ItemIndex = 1 then begin
      mode := iprod_mode;
      if GetProduction(mode,id,pname) = mrOk then
        name := string(pname)
      else
        raise Exception.Create('Не выбрано наименование продукции');
    end;

    // производим вызов сервисной процедуры поиска
    // по наименованию доп.расходов
    if conditionRadioGroup.ItemIndex = 2 then begin
      mode := iextra_item_mode;
      if GetProduction(mode,id,pname) = mrOk then
        name := string(pname)
      else
        raise Exception.Create('Не выбрано наименование доп.расходов ');
    end;

    if (name <> '') and (id <> 0) then begin
      // добавляем выбранную продукцию в список
      try
        conditionValueList.Add(IntToStr(id));
        conditionListBox.Items.Add(name);
      except
        on EStringListError do
        MessageDlg('Данный элемент уже содержится в списке', mtError,
                   [mbOk], 0);
      end;
    end;

  finally
    // разрешаем нажатие кнопки формирования отчета если список не пуст
    if conditionListBox.Items.Count <> 0 then
      sbReportToExcel.Enabled := true;
  end;
end;

end.
