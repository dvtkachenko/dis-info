unit invoicesUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin;

const
  iEnterprPage = 0;
  smakInvoicesTemplate = 'mak_invoices.xlt';

type
  TmakInvExportForm = class(TForm)
    InvPageControl: TPageControl;
    allInvQuery: TQuery;
    InvoiceItemsQuery: TQuery;
    ExtraInvoiceItemsQuery: TQuery;
    InvBeginMaskEdit: TMaskEdit;
    InvEndMaskEdit: TMaskEdit;
    forEnterprTabSheet: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    SkidkiPriplQuery: TQuery;
    acceptCheckBox: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure NotDetailInvoicesReport(Sender: TObject);
    procedure DetailInvoicesReport(Sender: TObject);
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

uses shared_type, excel_type;

{$R *.DFM}

function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';

{сервисные процедуры}

{-------------------}

procedure TmakInvExportForm.FormShow(Sender: TObject);
begin
  InvBeginMaskEdit.Text := startDate;
  InvEndMaskEdit.Text := DateToStr(Date);
end;


//---------------------------------------------------------------------
// выбор счетов-фактур закрепленных за отделами (не детализированный отчет)
//---------------------------------------------------------------------

procedure TmakInvExportForm.NotDetailInvoicesReport(Sender: TObject);
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
procedure TmakInvExportForm.DetailInvoicesReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     s_row : string;
     info_row : array[1..13] of Variant;
     PathToTemplate : string;
     i : integer;
     ItemsFlag, ExtraItemsFlag : boolean;
//     ReportHeader : string;
     dept_id : integer;
     row : integer;
     rowDetail : integer;

     { контрольные переменные }
     countInvoices : integer ;

     // invoice master
     invoice_id : integer;
     invoice_no : string;
     invoice_date : TDate;
     trade_mark : string;
     dimention : string;
     qnty : real;
     sum_without_nds : real;
     skidki_pripl : real;
     nds : real;
     price_without_nds : real;
     amount : real;
     cargo_sender_name : string;
     cargo_receiver_name : string;
     cargo_date : TDate;

//     full_sum : real;  // for extra_items

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + smakInvoicesTemplate;
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
     countInvoices := 0;
     row := 6;

     { просим в базе необходимые счета }
     allInvQuery.Open;
     InvoiceItemsQuery.Prepare;
     ExtraInvoiceItemsQuery.Prepare;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allInvQuery.Eof do begin
       countInvoices := countInvoices + 1;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

       // master
       invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
       with SkidkiPriplQuery do begin
         Close;
         ParamByName('invoice_id').asinteger := invoice_id;
         Open;
       end;
       invoice_no := allInvQuery.fieldbyname('invoice_no').asstring;
       invoice_no := ' ' + invoice_no;
       invoice_date := allInvQuery.fieldbyname('invoice_date').asdatetime;
       skidki_pripl := SkidkiPriplQuery.fieldbyname('skidki_pripl').asfloat;
       nds := allInvQuery.fieldbyname('nds').asfloat;
       amount := allInvQuery.fieldbyname('amount').asfloat;
       cargo_sender_name := allInvQuery.fieldbyname('cargo_sender').asstring;
       cargo_receiver_name := allInvQuery.fieldbyname('cargo_receiver').asstring;
       cargo_date := allInvQuery.fieldbyname('cargo_date').asdatetime;

       info_row[1] := invoice_no;
       info_row[2] := invoice_date;
       info_row[7] := skidki_pripl;
       info_row[8] := nds;
       info_row[10] := amount;
       info_row[11] := cargo_sender_name;
       info_row[12] := cargo_receiver_name;
       info_row[13] := cargo_date;


       cellFrom := 'A' + IntToStr(row);
       cellTo := 'M' + IntToStr(row);
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
         sum_without_nds := InvoiceItemsQuery.fieldbyname('summ_without_nds').asfloat;
         //
         info_row[3] := trade_mark;
         info_row[4] := dimention;
         info_row[5] := qnty;
         info_row[6] := sum_without_nds;
         info_row[9] := price_without_nds;

         // добавляем грузополучателей для каждой строки invoice_items
         info_row[2] := invoice_date;
         info_row[12] := cargo_receiver_name;
         info_row[13] := cargo_date;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'M' + IntToStr(row);
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         ItemsFlag := true;
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 13 do info_row[i] := unAssigned;

         row := row + 1;
         InvoiceItemsQuery.Next;
       end;

       ExtraItemsFlag := false;
       // extra invoice items
       while not ExtraInvoiceItemsQuery.eof do begin
         trade_mark := ExtraInvoiceItemsQuery.fieldbyname('extra_item_name').asstring;
         price_without_nds := ExtraInvoiceItemsQuery.fieldbyname('price_without_nds').asfloat;
//         full_sum := ExtraInvoiceItemsQuery.fieldbyname('full_price').asfloat;
         //
         info_row[3] := trade_mark;
         info_row[6] := price_without_nds;
         // добавляем грузополучателей для каждой строки invoice_items
         info_row[12] := cargo_receiver_name;
         info_row[13] := cargo_date;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'M' + IntToStr(row);
//         Excel.xla.Range[cellFrom,cellTo].Value := VarArrayOf(info_row);
         ExtraItemsFlag := true;
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 13 do info_row[i] := unAssigned;

         row := row + 1;
         ExtraInvoiceItemsQuery.Next;
       end;

       if (ItemsFlag = false) and (ExtraItemsFlag = false) then begin
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         row := row + 1;
       end;

       row := row + 1;
       allInvQuery.Next;
     end;

   finally
     Excel.free;
     allInvQuery.Close;
     ExtraInvoiceItemsQuery.Close;
     InvoiceItemsQuery.Close;
     SkidkiPriplQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

procedure TmakInvExportForm.sbReportToExcelClick(Sender: TObject);
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

  case InvPageControl.ActivePage.TabIndex of

    iEnterprPage :
       begin
         ReportHeader := 'Входящие ';
         with allInvQuery do begin
           Close;
           SQL.Clear;
           SQL.Add('select * from balans_report_all_invoices(:begin_date, :end_date)');
           SQL.Add('where sender_id = :id and payer_id = 0');
           SQL.Add('and is_in_oper = ''Y''');
           SQL.Add('order by sender_name, pay_date');
           Prepare;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end;

         if GetEnterprise(id,pname) = mrOk then begin
           name := string(pname);
           allInvQuery.ParamByName('id').asinteger := id;
         end
         else
          raise Exception.Create('Предприятие не выбрано');

       end; // конец iEnterprPage
  end;  // end of CASE


  ReportHeader := ReportHeader + 'счета-фактуры за период с ' +
                  InvBeginMaskEdit.Text + ' по ' + InvEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
  DetailInvoicesReport(Sender);
  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TmakInvExportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
