unit invoices_relUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, excel_type;

const
  iInvRelEnterprPage = 0;
  iInvRelDeptPage = 1;
  iInvRelCoalCoxPage = 2;
  sInvRelEnterprPage = 'InvRelEnterprTabSheet';
  sInvRelDeptPage = 'InvRelDeptTabSheet';
  sInvRelCoalCoxPage = 'InvRelCoalCoxTabSheet';
  sInvoicesTemplate = 'invoices_detail.xlt';

  // запрос на все входящие счета по отделу за указываемый период
  sDeptInInvSelect = 'select * from balans_report_all_invoices(:begin_date, :end_date)'
             + ' where dept_id = :id and payer_id = 0'
             + ' order by sender_name, pay_date';

  // запрос на все исходящие счета по отделу за указываемый период
  sDeptOutInvSelect = 'select * from balans_report_all_invoices(:begin_date, :end_date)'
             + ' where dept_id = :id and sender_id = 0'
             + ' order by payer_name, pay_date';

  // запрос на все входящие счета по предприятию за указываемый период
  sEnterprInInvSelect = 'select * from balans_report_all_invoices(:begin_date, :end_date)'
             + ' where sender_id = :id and payer_id = 0'
             + ' order by pay_date';

  // запрос на все исходящие счета по предприятию за указываемый период
  sEnterprOutInvSelect = 'select * from balans_report_all_invoices(:begin_date, :end_date)'
             + ' where payer_id = :id and sender_id = 0'
             + ' order by pay_date';

type
  TInvRelExportForm = class(TForm)
    InvPageControl: TPageControl;
    allInvQuery: TQuery;
    InvRelDeptTabSheet: TTabSheet;
    InvoiceItemsQuery: TQuery;
    ExtraInvoiceItemsQuery: TQuery;
    InvBeginMaskEdit: TMaskEdit;
    InvEndMaskEdit: TMaskEdit;
    InOutRadioGroup: TRadioGroup;
    InvRelEnterprTabSheet: TTabSheet;
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
    InvRelCoalCoxTabSheet: TTabSheet;
    Label8: TLabel;
    Label9: TLabel;
    SkidkiPriplQuery: TQuery;
    ccRadioGroup: TRadioGroup;
    SkidkiPriplCheckBox: TCheckBox;
    InvQuery: TQuery;
    zdtarifCheckBox: TCheckBox;
    zdtarifQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure InvoiceToExcel(Excel : TExcel;Var row : integer; countInv,invoice_id : integer);
    procedure InvRelReport(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure InvRelCoalCoxTabSheetShow(Sender: TObject);
    procedure InvRelCoalCoxTabSheetHide(Sender: TObject);
    procedure zdtarifCheckBoxClick(Sender: TObject);
    procedure SkidkiPriplCheckBoxClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    ReportHeader : string;
    BeginDate : TDateTime;
    EndDate : TDateTime;
    PathToProgram : string;
    sInvSlaveQuery : string;
    id : integer; // в зависимости от выбранной закладки есть
                  // или ent_id или dept_id
  end;

implementation

uses shared_type;

{$R *.DFM}

function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';

{сервисные процедуры}

{-------------------}

procedure TInvRelExportForm.FormShow(Sender: TObject);
begin
  InvBeginMaskEdit.Text := startDate;
  InvEndMaskEdit.Text := DateToStr(Date);
end;

//---------------------------------------------------------------------
// экспорт в Excel указанного счета-фактуры с указанной строки
// (детализированный отчет)
//---------------------------------------------------------------------
procedure TInvRelExportForm.InvoiceToExcel(Excel : TExcel; Var row : integer; countInv, invoice_id : integer);
  Var
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..34] of Variant;
     PathToTemplate : string;
     i : integer;
     ItemsFlag, ExtraItemsFlag : boolean;
//     ReportHeader : string;
     rowChain : integer;
     countInvSubItem : integer ;
     countChainInv : integer ;

     // invoice master
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
     nds : real;
     contract_no : string;
     cargo_sender_name : string;
     cargo_receiver_name : string;
     accept : string;
     reference : string;
     act_no : string;
     act_date : TDate;
     inv_type : string;
     dept_name : string;
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

  begin

   try

     rowChain := row; // запоминаем значение строки для связанных счетов
     // переменная для подсчета элементов в счете
     countInvSubItem := 0;

     Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
     // master
//     invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
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
     act_no := allInvQuery.fieldbyname('act_no').asstring;
     act_date := allInvQuery.fieldbyname('act_date').asdatetime;
     inv_type := allInvQuery.fieldbyname('programm_type').asstring;
     if inv_type = 'D' then inv_type := 'дав';
     if inv_type = 'U' then inv_type := 'обычн';
     if inv_type = 'E' then inv_type := 'экспорт';
     reference := allInvQuery.fieldbyname('is_in_ref').asstring;

     //  формируем столбец приплаты/скидки
     if (InvPageControl.ActivePage.TabIndex = iInvRelCoalCoxPage) and
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
     end;

     info_row[1] := countInv;
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
       ItemsFlag := true;
       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       // красим в желтый цвет неакцептованные счета
       if (accept = 'N') then Excel.FillRangeColor(cellFrom,cellTo,6);

       for i := 1 to 34 do info_row[i] := unAssigned;

       row := row + 1;
       InvoiceItemsQuery.Next;
     end;

     ExtraItemsFlag := false;
     // extra invoice items
     while not ExtraInvoiceItemsQuery.eof do begin
       trade_mark := ExtraInvoiceItemsQuery.fieldbyname('extra_item_name').asstring;
       price_without_nds := ExtraInvoiceItemsQuery.fieldbyname('price_without_nds').asfloat;
       full_price := ExtraInvoiceItemsQuery.fieldbyname('full_price').asfloat;
       //
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
       ExtraItemsFlag := true;
       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       // красим в желтый цвет неакцептованные счета
       if (accept = 'N') then Excel.FillRangeColor(cellFrom,cellTo,6);

       for i := 1 to 34 do info_row[i] := unAssigned;

       row := row + 1;
       ExtraInvoiceItemsQuery.Next;
     end;

     if (ItemsFlag = false) and (ExtraItemsFlag = false) then begin
         countInvSubItem := countInvSubItem + 1;
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         // красим в желтый цвет неакцептованные счета
         if (accept = 'N') then Excel.FillRangeColor(cellFrom,cellTo,6);
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

         // увеличиваем счетчик привязанных счетов на 1
         countChainInv := countChainInv + 1;
         info_row[1] := c_invoice_no;
         info_row[2] := c_pay_date;
         info_row[3] := c_invoice_date;
         info_row[4] := c_contract_no;
         info_row[5] := c_sender_name;
         info_row[6] := c_payer_name;

         cellFrom := 'AC' + IntToStr(rowChain);
         cellTo := 'AH' + IntToStr(rowChain);
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
           Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

           rowChain := rowChain + 1;
         end;
       end;

       if row < rowChain then row := rowChain;

     end;

   finally
     ExtraInvoiceItemsQuery.Close;
     InvoiceItemsQuery.Close;
     zdtarifQuery.Close;
   end;
end;

//---------------------------------------------------------------------
// выбор связанных счетов-фактур и формирование отчета
// (детализированный отчет)
//---------------------------------------------------------------------
procedure TInvRelExportForm.InvRelReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     PathToTemplate : string;
     row : integer;
     { контрольные переменные }
     countInvoices : integer ;

     // invoice master
     invoice_id : integer;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sInvoicesTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try
     // копируем лист для связанных счетов
     Excel.CopyWorkSheet('invoices','rel_invoices');

     ///////////////////////////////////////////////////////////
     // формирование отчета по связанным счетам-фактурам MASTER
     ///////////////////////////////////////////////////////////
     Excel.SelectWorkSheet('invoices');
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     if (InvPageControl.ActivePage.TabIndex = iInvRelCoalCoxPage) and
        (SkidkiPriplCheckBox.Checked) then begin
       cell := 'Q' + IntToStr(4);
       Excel.Cell[cell] := 'Приплаты/Скидки без НДС';
     end;

     chainCheckBox.Checked := true;
     // формируем запрос для связанных с MASTER-счетами счетов
     if chainCheckBox.Checked then begin
       with chainInvQuery do begin
         Close;
         SQL.Clear;
         SQL.Add('select i.invoice_id, i.invoice_no, i.rate_date pay_date, i.invoice_date,');
         SQL.Add('s.enterprise_name sender_name, p.enterprise_name payer_name,');
         SQL.Add('i.contract_no');
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

     // если выбран "показать ж/д тариф" , то меняем соответствующим образом
     // заголовок столбца в отчете
     if zdtarifCheckBox.Checked then begin
       cell := 'Q' + IntToStr(4);
       Excel.Cell[cell] := 'Ж/д тариф по квитанции б/НДС';
     end;

     { инициализируем  контрольные переменные }
     countInvoices := 0;
     row := 6;

     { просим в базе необходимые счета }
     allInvQuery.Open;
     InvoiceItemsQuery.Prepare;
     ExtraInvoiceItemsQuery.Prepare;

  // ---- ---- ----- начало цикла по счетам  MASTER ----- ----- ----- //
     while not allInvQuery.Eof do begin
       countInvoices := countInvoices + 1;
       // экспортируем в Excel счет - фактуру
       invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
       InvoiceToExcel(Excel,row,countInvoices,invoice_id);
       row := row + 1;
       allInvQuery.Next;
     end;

     ///////////////////////////////////////////////////////////
     // формирование отчета по связанным счетам-фактурам SLAVE
     ///////////////////////////////////////////////////////////
     Excel.SelectWorkSheet('rel_invoices');
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := 'Связанные счета - SLAVE';

     if (InvPageControl.ActivePage.TabIndex = iInvRelCoalCoxPage) and
        (SkidkiPriplCheckBox.Checked) then begin
       cell := 'Q' + IntToStr(4);
       Excel.Cell[cell] := 'Приплаты/Скидки без НДС';
     end;

     chainCheckBox.Checked := true;
     // формируем запрос для связанных со SLAVE-счетами счетов
     if chainCheckBox.Checked then begin
       with chainInvQuery do begin
         Close;
         SQL.Clear;
         SQL.Add('select i.invoice_id, i.invoice_no, i.rate_date pay_date, i.invoice_date,');
         SQL.Add('s.enterprise_name sender_name, p.enterprise_name payer_name,');
         SQL.Add('i.contract_no');
         SQL.Add('from invoices i,');
         SQL.Add('invoice_references ir,');
         SQL.Add('enterpr s,');
         SQL.Add('enterpr p');
         if InOutRadioGroup.ItemIndex = 0 then begin
           SQL.Add('where i.invoice_id = ir.in_id');
           SQL.Add('and ir.out_id = :inv_id');
         end;
         if InOutRadioGroup.ItemIndex = 1 then begin
           SQL.Add('where i.invoice_id = ir.out_id');
           SQL.Add('and ir.in_id = :inv_id');
         end;
         SQL.Add('and i.sender_id = s.enterpr_id');
         SQL.Add('and i.payer_id = p.enterpr_id');
         SQL.Add('order by i.rate_date');
         Prepare;
       end  // end of WITH
     end;   // end of IF

     // если выбран "показать ж/д тариф" , то меняем соответствующим образом
     // заголовок столбца в отчете
     if zdtarifCheckBox.Checked then begin
       cell := 'Q' + IntToStr(4);
       Excel.Cell[cell] := 'Ж/д тариф по квитанции б/НДС';
     end;

     { инициализируем  контрольные переменные }
     countInvoices := 0;
     row := 6;

     with allInvQuery do begin
       Close;
       SQL.Clear;
       SQL.Add(sInvSlaveQuery);
       Prepare;
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
//       ParamByName('id').asdate := id;
     end;
     allInvQuery.Open;
  // ---- ---- ----- начало цикла по счетам  SLAVE ----- ----- ----- //
     while not allInvQuery.Eof do begin
       // экспортируем в Excel счет - фактуру
       countInvoices := countInvoices + 1;
       invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
       InvoiceToExcel(Excel,row,countInvoices,invoice_id);
       row := row + 1;
       allInvQuery.Next;
     end;

   finally
     Excel.free;
     chainCheckBox.Checked := true;
     allInvQuery.Close;
     ExtraInvoiceItemsQuery.Close;
     InvoiceItemsQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

// ---------------------------------------------------------------
procedure TInvRelExportForm.sbReportToExcelClick(Sender: TObject);
Var
//  id : integer;
  name : string;
  s : array[0..maxPChar] of Char;
  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
  pname := @s;
  BeginDate := StrToDate(InvBeginMaskEdit.Text);
  EndDate := StrToDate(InvEndMaskEdit.Text);

  case InvPageControl.ActivePage.TabIndex of

    iInvRelDeptPage :
       begin
         if GetDepatment(id,pname) = mrOk then begin
           name := string(pname);
         end
         else
          raise Exception.Create('Отдел не выбран');

         if InOutRadioGroup.ItemIndex = 0 then begin
           ReportHeader := 'Входящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add(sDeptInInvSelect);
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
             ParamByName('id').asinteger := id;
           end;
           sInvSlaveQuery := 'select distinct io.* from balans_invoices_list(1) io,'
                             + ' balans_report_all_invoices(:begin_date, :end_date) ii,'
                             + ' invoice_references ir'
                             + ' where io.invoice_id = ir.out_id'
                             + ' and ii.invoice_id = ir.in_id'
                             + ' and ii.dept_id = '
                             + IntToStr(id)
                             + ' and ii.payer_id = 0'
                             + ' order by io.payer_id, io.pay_date';
         end;

         if InOutRadioGroup.ItemIndex = 1 then begin
           ReportHeader := 'Исходящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add(sDeptOutInvSelect);
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
             ParamByName('id').asinteger := id;
           end;
           sInvSlaveQuery := 'select distinct ii.* from balans_invoices_list(1) ii,'
                             + ' balans_report_all_invoices(:begin_date, :end_date) io,'
                             + ' invoice_references ir'
                             + ' where io.invoice_id = ir.out_id'
                             + ' and ii.invoice_id = ir.in_id'
                             + ' and io.dept_id = '
                             + IntToStr(id)
                             + ' and io.sender_id = 0'
                             + ' order by ii.sender_id, ii.pay_date';
         end;

         ReportHeader := ReportHeader + 'счета-фактуры за период с ' +
                  InvBeginMaskEdit.Text + ' по ' + InvEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
         InvRelReport(Sender);

       end; // конец iDeptPage

    iInvRelEnterprPage :
       begin
         if GetEnterprise(id,pname) = mrOk then begin
           name := string(pname);
         end
         else
          raise Exception.Create('Предприятие не выбрано');

         if InOutRadioGroup.ItemIndex = 0 then begin
           ReportHeader := 'Входящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add(sEnterprInInvSelect);
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
             ParamByName('id').asinteger := id;
           end;
           sInvSlaveQuery := 'select distinct io.* from balans_invoices_list(1) io,'
                             + ' balans_report_all_invoices(:begin_date, :end_date) ii,'
                             + ' invoice_references ir'
                             + ' where io.invoice_id = ir.out_id'
                             + ' and ii.invoice_id = ir.in_id'
                             + ' and ii.sender_id = '
                             + IntToStr(id)
                             + ' and ii.payer_id = 0'
                             + ' order by io.payer_id, io.pay_date';
         end;

         if InOutRadioGroup.ItemIndex = 1 then begin
           ReportHeader := 'Исходящие ';
           with allInvQuery do begin
             Close;
             SQL.Clear;
             SQL.Add(sEnterprOutInvSelect);
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
             ParamByName('id').asinteger := id;
           end;
{           sInvSlaveQuery := 'select distinct ii.* from balans_invoices_list(1) ii,'
                             + ' balans_report_all_invoices(:begin_date, :end_date) io,'
                             + ' invoice_references ir'
                             + ' where io.invoice_id = ir.out_id'
                             + ' and ii.invoice_id = ir.in_id'
                             + ' and io.payer_id = :id and io.sender_id = 0'
                             + ' order by ii.payer_id, ii.pay_date'; }
           sInvSlaveQuery := 'select distinct ii.* from balans_invoices_list(1) ii,'
                             + ' balans_report_all_invoices(:begin_date, :end_date) io,'
                             + ' invoice_references ir'
                             + ' where io.invoice_id = ir.out_id'
                             + ' and ii.invoice_id = ir.in_id'
                             + ' and io.payer_id = '
                             + IntToStr(id)
                             + ' and io.sender_id = 0'
                             + ' order by ii.sender_id, ii.pay_date';
         end;

         ReportHeader := ReportHeader + 'счета-фактуры за период с ' +
                  InvBeginMaskEdit.Text + ' по ' + InvEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
         InvRelReport(Sender);

       end; // конец iEnterprPage

{    iInvRelCoalCoxPage :
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
             if ccRadioGroup.ItemIndex = 0 then
               SQL.Add('(s.prod_id = 10010 or s.prod_id = 10011)');
             if ccRadioGroup.ItemIndex = 1 then
               SQL.Add('(s.prod_id = 13)');
             SQL.Add('order by sender_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end;
           name := 'уголь';
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
             if ccRadioGroup.ItemIndex = 0 then
               SQL.Add('(s.prod_id = 10010 or s.prod_id = 10011)');
             if ccRadioGroup.ItemIndex = 1 then
               SQL.Add('(s.prod_id = 13)');
             SQL.Add('order by sender_name, pay_date');
             Prepare;
             ParamByName('begin_date').asdate := BeginDate;
             ParamByName('end_date').asdate := EndDate;
           end;
           name := 'кокс';
         end;
         ReportHeader := ReportHeader + 'счета-фактуры за период с ' +
                  InvBeginMaskEdit.Text + ' по ' + InvEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
         InvRelReport(Sender);

       end; // конец iInvRelCoalCoxPage }
  end;  // end of CASE

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TInvRelExportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TInvRelExportForm.InvRelCoalCoxTabSheetShow(Sender: TObject);
begin
  SkidkiPriplCheckBox.Enabled := true;
  SkidkiPriplCheckBox.Visible := true;
end;

procedure TInvRelExportForm.InvRelCoalCoxTabSheetHide(Sender: TObject);
begin
  SkidkiPriplCheckBox.Checked := false;
  SkidkiPriplCheckBox.Enabled := false;
  SkidkiPriplCheckBox.Visible := false;
end;

procedure TInvRelExportForm.zdtarifCheckBoxClick(Sender: TObject);
begin
  if SkidkiPriplCheckBox.Checked then begin
    zdtarifCheckBox.Checked := false;
  end;
end;

procedure TInvRelExportForm.SkidkiPriplCheckBoxClick(Sender: TObject);
begin
  if zdtarifCheckBox.Checked then begin
    SkidkiPriplCheckBox.Checked := false;
  end;
end;

end.
