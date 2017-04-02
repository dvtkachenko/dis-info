unit contract_statisticUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, DBCtrls, shared_type;

const
  schoose_contract = 'choose_contractTabSheet';
  sstatistic_contract = 'statistic_contractTabSheet';
  sreportChangePage = 'changeTabSheet';
  scontract_grpTemplate = 'statistic.xlt';
//  sChangeTemplate = 'change_tcredit.xlt';
type

  TContract_grp = class (TObject)
  private
    FcntrQuery : TQuery;
    FcntrComboBox : TComboBox;
    // массив для иденификаторов групп договоров
    Frel_type_id : array [0..imaxCntrItem] of real;
    Fent_type_id : array [0..imaxCntrItem] of real;
    // реальное кол-во элементов в массиве
    FcountItem : integer;
    Findex : integer;
    Frel_name : string;
  public
    constructor Create(cntrQuery : TQuery;
                       cntrComboBox : TComboBox);
    function GetCurRelTypeID : real;
    function GetCurEntTypeID : real;
    destructor Destroy; override;
  end;

  Tcontract_statisticForm = class(TForm)
    contract_grpPageControl: TPageControl;
    allContractQuery: TQuery;
    statistic_contractTabSheet: TTabSheet;
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
    contract_saldoQuery: TQuery;
    check_oper_contractQuery: TQuery;
    choose_contractTabSheet: TTabSheet;
    reportBeginMaskEdit: TMaskEdit;
    reportEndMaskEdit: TMaskEdit;
    contractDBGrid: TDBGrid;
    contractDataSource: TDataSource;
    contractAddBitBtn: TBitBtn;
    contractDelBitBtn: TBitBtn;
    insertContractQuery: TQuery;
    deleteContractQuery: TQuery;
    changeTabSheet: TTabSheet;
    Label3: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    journalDateMaskEdit: TMaskEdit;
    Label9: TLabel;
    changeBeginMaskEdit: TMaskEdit;
    changeEndMaskEdit: TMaskEdit;
    changeTCQuery: TQuery;
    disContractInvoiceInQuery: TQuery;
    disContractInvoiceOutQuery: TQuery;
    disContractPlategiQueryDebitor: TQuery;
    disContractPlategiQueryCreditor: TQuery;
    disContractAnyQueryDebitor: TQuery;
    disContractAnyQueryCreditor: TQuery;
    checkContractOperationQuery: TQuery;
    contract_grpQuery: TQuery;
    contract_grpDataSource: TDataSource;
    contract_grpComboBox: TComboBox;
    allContractGridQuery: TQuery;
    readonlyCheckBox: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure exportReport(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure contractAddBitBtnClick(Sender: TObject);
    procedure contractDelBitBtnClick(Sender: TObject);
    procedure exportChange(Sender: TObject);
    procedure contract_grpComboBoxChange(Sender: TObject);
    procedure readonlyCheckBoxClick(Sender: TObject);
    procedure statistic_contractTabSheetEnter(Sender: TObject);
    procedure choose_contractTabSheetEnter(Sender: TObject);
    procedure changeTabSheetEnter(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    cntr_grp : TContract_grp;
    ReportHeader : string;
    BeginDate : TDateTime;
    EndDate : TDateTime;
    PathToProgram : string;
  end;

implementation

uses excel_type;

{$R *.DFM}

function GetDepatment(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetDepatment';
function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';
function GetContract(id:integer;Var contract_id:integer;Var pcontract_no: PChar) : integer; external 'service.dll' name 'GetContract';


{сервисные процедуры}

{-------------------}

// ---------------------------------------
// --- реализация методов класса TContract_grp
// ---------------------------------------
constructor TContract_grp.Create(cntrQuery : TQuery;
                                 cntrComboBox : TComboBox);
begin
  inherited Create;
  FcountItem := 0;
  FcntrQuery := cntrQuery;
  FcntrComboBox := cntrComboBox;
  // заполняем элементы выпадающего списка типами договоров
  // по которым необходима статистика
  FcntrQuery.Close;
  FcntrQuery.Open;

  while not FcntrQuery.Eof do begin
    FcountItem := FcountItem + 1;
    if (FcountItem < imaxCntrItem) then begin
      Frel_name := FcntrQuery.fieldbyname('rel_name').asstring;
      Findex := FcntrComboBox.Items.Add(Frel_name);
      Frel_type_id[Findex] := FcntrQuery.fieldbyname('rel_type_id').asfloat;
      Fent_type_id[Findex] := FcntrQuery.fieldbyname('type_entity_id_1').asfloat;
    end;
    FcntrQuery.Next;
  end;

  FcntrQuery.Close;
end;
//---------------------------------------
function TContract_grp.GetCurRelTypeID : real;
begin
  Findex := FcntrComboBox.ItemIndex;
  Result := Frel_type_id[Findex];
end;
//---------------------------------------
function TContract_grp.GetCurEntTypeID : real;
begin
  Findex := FcntrComboBox.ItemIndex;
  Result := Fent_type_id[Findex];
end;
//---------------------------------------
destructor TContract_grp.Destroy;
begin
  FcntrQuery := nil;
  FcntrComboBox := nil;
  inherited Destroy;
end;

// ---------------------------------------
// --- реализация методов класса Tcontract_statisticForm
// ---------------------------------------
procedure Tcontract_statisticForm.FormShow(Sender: TObject);
begin
  reportBeginMaskEdit.Text := startDate;
  reportEndMaskEdit.Text := DateToStr(Date);
  changeBeginMaskEdit.Text := startDate;
  changeEndMaskEdit.Text := DateToStr(Date);
  journalDateMaskEdit.Text := DateToStr(Date-1);
  allContractQuery.Close;
  cntr_grp := TContract_grp.Create(contract_grpQuery, contract_grpComboBox);
  contract_grpComboBox.Text := '';
  // пока не выбрана группа договоров
  // блокируем дальнейшую работу
  sbReportToExcel.Enabled := false;
  statistic_contractTabSheet.Enabled := false;
  changeTabSheet.Enabled := false;
  contractAddBitBtn.Enabled := false;
  contractDelBitBtn.Enabled := false;
  readonlyCheckBox.Checked := true;
  readonlyCheckBox.Enabled := false;
end;

//---------------------------------------------------------------------
// формирует отчет о работе ДП ДИС по
// выбранным типам договоров
//---------------------------------------------------------------------
procedure Tcontract_statisticForm.exportReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..15] of Variant;
     PathToTemplate : string;

     BeginDate : TDateTime;
     EndDate : TDateTime;
     ReportHeader : string;
     ent_id : real;
     ent_name : string;
     rowDebit,rowCredit,row : integer;
     i : integer;

     contractSaldoBegin, contractSaldoEnd : real;
     contractDebitAccept, contractCreditAccept : real;
     contractDebitNoAccept, contractCreditNoAccept : real;
     contractSaldoAccept, contractSaldoNoAccept : real;

     { контрольные переменные }
     countContract : integer ;
     Contract : string;

     all_contractes_saldo : real;  // сальдо по всем договорам

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
     ColumnEnterpr = 'A';
     ColumnCntr = 'B';
     ColumnDebitStart = 'D';
     ColumnDebitEnd = 'O';
     ColumnName = 'P';
     ColumnCreditStart = 'Q';
     ColumnCreditEnd = 'AB';
     ColumnEntName = 'G';
     ColumnContract = 'H';
     ColumnSum = 'J';
     ColumnEnd = ColumnCreditEnd;

  begin
     temp := GetThreadLocale;
     SetThreadLocale(English_Locale);

     try
       Excel := TExcel.Create;
     except
       raise Exception.Create('Невозможно создать OLE - объект');
     end;

     PathToTemplate := PathToProgram + '\Template\' + scontract_grpTemplate;
     try
       Excel.AddWorkBook(PathToTemplate);
       Excel.Visible := true;
     except
       raise Exception.Create('Невозможно загрузить Excel');
     end;

  try
     ReportHeader := 'Отчет : ' +
                  contract_grpComboBox.Text + ' : ' +
                  reportBeginMaskEdit.Text + ' по ' + reportEndMaskEdit.Text
                  + ' (' + TimeToStr(Time) + ')';

     row := 2;
     cell := ColumnEnterpr + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     BeginDate := StrToDate(reportBeginMaskEdit.Text);
     EndDate := StrToDate(reportEndMaskEdit.Text);

     { инициализируем  контрольные переменные }
     countContract := 0;

     all_contractes_saldo := 0;

     row := 7;
     rowCredit := 7;
     rowDebit := 7;

     { формируем список всех контрактов выбранного типа}
     with allContractQuery do begin
       Close;
       ParamByName('rel_type_id').asfloat := cntr_grp.GetCurRelTypeID;
       ParamByName('ent_type_id').asfloat := cntr_grp.GetCurEntTypeID;
       Open;
     end;

     // ---- ---- ----- начало цикла по договорам ----- ----- ----- //
     while not allContractQuery.Eof do begin
       //
       ent_id := allContractQuery.fieldbyname('enterpr_id').asinteger;
       ent_name := allContractQuery.fieldbyname('enterprise_name').asstring;
       contract := allContractQuery.fieldbyname('contract_no').asstring;

       contractSaldoBegin := 0;
       with contract_saldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('contract_no').asstring := contract;
         { берем сальдо на день раньше }
         ParamByName('saldo_date').asdate := BeginDate - 1;
         Open;
       end;
       contractSaldoBegin := contract_saldoQuery.fieldbyname('debit').asfloat
                             - contract_saldoQuery.fieldbyname('credit').asfloat;

       contractSaldoEnd := 0;
       with contract_saldoQuery do begin
         Close;
         ParamByName('ent_id').asfloat := ent_id;
         ParamByName('contract_no').asstring := contract;
         ParamByName('saldo_date').asdate := EndDate;
         Open;
       end;
       contractSaldoEnd := contract_saldoQuery.fieldbyname('debit').asfloat
                             - contract_saldoQuery.fieldbyname('credit').asfloat;

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

           allContractQuery.Next;
           continue;
       end;

       // увеличиваем счетчик договоров 
       countContract := countContract + 1;

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

       cell := ColumnEnterpr + IntToStr(row);
       Excel.Cell[cell] :=
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------';

       cellFrom := ColumnEnterpr + IntToStr(row);
       cellTo := columnEnd + IntToStr(row);
       Excel.FillRangeColor(cellFrom, cellTo, 6);

       row := row + 2;
       cell := ColumnEntName + IntToStr(row);
       Excel.Cell[cell] := 'Сальдо на начало периода';

       row := row + 1;
       cell := ColumnEntName + IntToStr(row);
       Excel.Cell[cell] := 'Предприятие : ' + ent_name;

       row := row + 1;
       cell := ColumnEntName + IntToStr(row);
       Excel.Cell[cell] := 'Договор : ' + contract + ' :';

       cell := ColumnSum + IntToStr(row);
       Excel.Cell[cell] := contractSaldoBegin;

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
         info_row[12] := comment;

         cellFrom := columnDebitStart + IntToStr(rowDebit);
         cellTo := columnDebitEnd + IntToStr(rowDebit);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 12 do info_row[i] := unAssigned;

         //  для удобства фильтрации в Excel
         cell := ColumnEnterpr + IntToStr(rowDebit);
         Excel.Cell[cell] := ent_name;
         cell := ColumnCntr + IntToStr(rowDebit);
         Excel.Cell[cell] := contract;

         rowDebit := rowDebit + 1;
         disContractInvoiceInQuery.Next;
       end;
       disContractInvoiceInQuery.Close;
       rowDebit := rowDebit + 2;

       while not disContractPlategiQueryDebitor.Eof do begin
         pay_date := disContractPlategiQueryDebitor.fieldbyname('doc_date').asdatetime;
//         cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disContractPlategiQueryDebitor.fieldbyname('type_name').asstring;
         doc_no := disContractPlategiQueryDebitor.fieldbyname('pay_order').asstring;
//         short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disContractPlategiQueryDebitor.fieldbyname('amount').asfloat;
         amount_usd := disContractPlategiQueryDebitor.fieldbyname('amount_usd').asfloat;
         contract_no := disContractPlategiQueryDebitor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disContractPlategiQueryDebitor.fieldbyname('comment').asstring;
         if accept = 'Y' then contractDebitAccept := contractDebitAccept + Amount;
         contractDebitNoAccept := contractDebitNoAccept + Amount;

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

         cellFrom := columnDebitStart + IntToStr(rowDebit);
         cellTo := columnDebitEnd + IntToStr(rowDebit);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 12 do info_row[i] := unAssigned;

         //  для удобства фильтрации в Excel
         cell := ColumnEnterpr + IntToStr(rowDebit);
         Excel.Cell[cell] := ent_name;
         cell := ColumnCntr + IntToStr(rowDebit);
         Excel.Cell[cell] := contract;

         rowDebit := rowDebit + 1;
         disContractPlategiQueryDebitor.Next;
       end;
       disContractPlategiQueryDebitor.Close;
       rowDebit := rowDebit + 2;

       while not disContractAnyQueryDebitor.Eof do begin
         pay_date := disContractAnyQueryDebitor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disAnyQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disContractAnyQueryDebitor.fieldbyname('type_name').asstring;
//         doc_no := disAnyQueryDebitor.fieldbyname('act_no').asstring;
//         short_trade_mark := disAnyQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disContractAnyQueryDebitor.fieldbyname('amount').asfloat;
         amount_usd := disContractAnyQueryDebitor.fieldbyname('amount_usd').asfloat;
         contract_no := disContractAnyQueryDebitor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disContractAnyQueryDebitor.fieldbyname('comment').asstring;
         if accept = 'Y' then contractDebitAccept := contractDebitAccept + Amount;
         contractDebitNoAccept := contractDebitNoAccept + Amount;

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

         cellFrom := columnDebitStart + IntToStr(rowDebit);
         cellTo := columnDebitEnd + IntToStr(rowDebit);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 12 do info_row[i] := unAssigned;

         //  для удобства фильтрации в Excel
         cell := ColumnEnterpr + IntToStr(rowDebit);
         Excel.Cell[cell] := ent_name;
         cell := ColumnCntr + IntToStr(rowDebit);
         Excel.Cell[cell] := contract;

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
//       comment := ;
         if accept = 'Y' then contractCreditAccept := contractCreditAccept + Amount;
         contractCreditNoAccept := contractCreditNoAccept + Amount;

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
//         info_row[12] := comment;

         cellFrom := columnCreditStart + IntToStr(rowCredit);
         cellTo := columnCreditEnd + IntToStr(rowCredit);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 12 do info_row[i] := unAssigned;

         //  для удобства фильтрации в Excel
         cell := ColumnEnterpr + IntToStr(rowCredit);
         Excel.Cell[cell] := ent_name;
         cell := ColumnCntr + IntToStr(rowCredit);
         Excel.Cell[cell] := contract;

         rowCredit := rowCredit + 1;
         disContractInvoiceOutQuery.Next;
       end;
       disContractInvoiceOutQuery.Close;
       rowCredit := rowCredit + 2;

       { вся кредитовая статистика кроме тов.отгрузок }
       while not disContractPlategiQueryCreditor.Eof do begin
         pay_date := disContractPlategiQueryCreditor.fieldbyname('doc_date').asdatetime;
//         cargo_date := disPlategiQueryCreditor.fieldbyname('cargo_date').asdatetime;
         doc_type := disContractPlategiQueryCreditor.fieldbyname('type_name').asstring;
         doc_no := disContractPlategiQueryCreditor.fieldbyname('pay_order').asstring;
//       short_trade_mark := disPlategiQueryCreditor.fieldbyname('short_trade_mark').asstring;
         amount := disContractPlategiQueryCreditor.fieldbyname('amount').asfloat;
         amount_usd := disContractPlategiQueryCreditor.fieldbyname('amount_usd').asfloat;
         contract_no := disContractPlategiQueryCreditor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disContractPlategiQueryCreditor.fieldbyname('comment').asstring;
         if accept = 'Y' then contractCreditAccept := contractCreditAccept + Amount;
         contractCreditNoAccept := contractCreditNoAccept + Amount;

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

         cellFrom := columnCreditStart + IntToStr(rowCredit);
         cellTo := columnCreditEnd + IntToStr(rowCredit);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 12 do info_row[i] := unAssigned;

         //  для удобства фильтрации в Excel
         cell := ColumnEnterpr + IntToStr(rowCredit);
         Excel.Cell[cell] := ent_name;
         cell := ColumnCntr + IntToStr(rowCredit);
         Excel.Cell[cell] := contract;

         rowCredit := rowCredit + 1;
         disContractPlategiQueryCreditor.Next;
       end;
       disContractPlategiQueryCreditor.Close;
       rowCredit := rowCredit + 2;

       while not disContractAnyQueryCreditor.Eof do begin
         pay_date := disContractAnyQueryCreditor.fieldbyname('pay_date').asdatetime;
//         cargo_date := disPlategiQueryDebitor.fieldbyname('cargo_date').asdatetime;
         doc_type := disContractAnyQueryCreditor.fieldbyname('type_name').asstring;
//         doc_no := disAnyQueryCreditor.fieldbyname('act_no').asstring;
//         short_trade_mark := disPlategiQueryDebitor.fieldbyname('short_trade_mark').asstring;
         amount := disContractAnyQueryCreditor.fieldbyname('amount').asfloat;
         amount_usd := disContractAnyQueryCreditor.fieldbyname('amount_usd').asfloat;
         contract_no := disContractAnyQueryCreditor.fieldbyname('contract_no').asstring;
         accept := 'Y';
         comment := disContractAnyQueryCreditor.fieldbyname('comment').asstring;
         if accept = 'Y' then contractCreditAccept := contractCreditAccept + Amount;
         contractCreditNoAccept := contractCreditNoAccept + Amount;

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

         cellFrom := columnCreditStart + IntToStr(rowCredit);
         cellTo := columnCreditEnd + IntToStr(rowCredit);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
         for i := 1 to 12 do info_row[i] := unAssigned;

         //  для удобства фильтрации в Excel
         cell := ColumnEnterpr + IntToStr(rowCredit);
         Excel.Cell[cell] := ent_name;
         cell := ColumnCntr + IntToStr(rowCredit);
         Excel.Cell[cell] := contract;

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
       cell := ColumnEntName + IntToStr(row);
       Excel.Cell[cell] := 'Сальдо на конец периода';
       cell := ColumnSum + IntToStr(row);
       Excel.Cell[cell] := contractSaldoEnd;

       row := row + 1;
       cell := ColumnEntName + IntToStr(row);
       Excel.Cell[cell] := 'сальдо c акцептом:';
       cell := ColumnSum + IntToStr(row);
       Excel.Cell[cell] := contractSaldoAccept;

       row := row + 1;
       cell := ColumnEntName + IntToStr(row);
       Excel.Cell[cell] := 'сальдо без акцепта:';
       cell := ColumnSum + IntToStr(row);
       Excel.Cell[cell] := contractSaldoNoAccept;

       row := row + 2;

       allContractQuery.Next;
       Update;
     end; // конец    while not allContractQuery.Eof

     cell := ColumnEnterpr + IntToStr(row);
     Excel.Cell[cell] :=
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------' +
           '-----------------------------------------------------------';

     cellFrom := ColumnEnterpr + IntToStr(row);
     cellTo := columnEnd + IntToStr(row);
     Excel.FillRangeColor(cellFrom, cellTo, 6);

     row := row + 2;
     cell := ColumnEntName + IntToStr(row);
     Excel.Cell[cell] := 'СУММА САЛЬДО С АКЦЕПТОМ ПО ВСЕМ ДОГОВОРАМ:';
     cell := ColumnSum + IntToStr(row);
     Excel.Cell[cell] := all_contractes_saldo;

     row := row + 1;
     cell := ColumnEntName + IntToStr(row);
     Excel.Cell[cell] := 'Общее количество договоров:';
     cell := ColumnSum + IntToStr(row);
     Excel.Cell[cell] := countContract;

  finally
//      allSaldoQuery.Close;
    allContractQuery.Close;
    contract_saldoQuery.Close;
    checkContractOperationQuery.Close;
    disContractPlategiQueryDebitor.Close;
    disContractPlategiQueryCreditor.Close;
    disContractAnyQueryDebitor.Close;
    disContractAnyQueryCreditor.Close;
    disContractInvoiceOutQuery.Close;
    disContractInvoiceInQuery.Close;
    Excel.free;
    SetThreadLocale(Temp);
  end;
end;

//---------------------------------------------------------------------
// формирование отчета об изменениях в базе за отчетный период
// на определенную дату по выбранной группе договоров
//---------------------------------------------------------------------
procedure Tcontract_statisticForm.exportChange(Sender: TObject);
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
//  PathToTemplate := PathToProgram + '\Template\' + sChangeTemplate;
  try
    Excel.AddWorkBook(PathToTemplate);
    Excel.Visible := true;
  except
    raise Exception.Create('Невозможно загрузить Excel');
  end;

  try
    ReportHeader := 'Журнал изменений в базе данных ДИСа '
                      + 'по договорам товарного кредита'
                      + ' за отчетный период с '
                      + changeBeginMaskEdit.Text
                      + ' по '
                      + changeEndMaskEdit.Text
                      + ' начиная с '
                      + JournalDateMaskEdit.Text
                      + ' (' + TimeToStr(Time) + ')';

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
procedure Tcontract_statisticForm.sbReportToExcelClick(Sender: TObject);
begin
  { конструирование запросов }
  BeginDate := StrToDate(reportBeginMaskEdit.Text);
  EndDate := StrToDate(reportEndMaskEdit.Text);

  if contract_grpPageControl.ActivePage.Name = sstatistic_contract then
       begin
         // формируем отчет
         exportReport(Sender);
       end; // конец sstatistic_contract

  if contract_grpPageControl.ActivePage.Name = sreportChangePage then
       begin
         // формируем отчет
         exportChange(Sender);
       end; // конец sreportChangePage

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure Tcontract_statisticForm.ExitSpeedButtonClick(Sender: TObject);
begin
  cntr_grp.Free;
  allContractGridQuery.Close;
  Close;
end;

//---------------------------------------------------------------
// процедура добавления договора
//---------------------------------------------------------------
procedure Tcontract_statisticForm.contractAddBitBtnClick(Sender: TObject);
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
      with insertContractQuery do begin
        Close;
        ParamByName('contract_id').asinteger:= contract_id;
        ParamByName('rel_type_id').asfloat := cntr_grp.GetCurRelTypeID;
        ParamByName('ent_type_id').asfloat := cntr_grp.GetCurEntTypeID;
      end;
      insertContractQuery.ExecSQL;
      // перечитываем данные
      allContractGridQuery.Close;
      allContractGridQuery.Open;
    end
    else
      raise Exception.Create('Договор не выбран');
  end
  else
    raise Exception.Create('Предприятие не выбрано');
end;

//---------------------------------------------------------------
// процедура удаления договора
//---------------------------------------------------------------
procedure Tcontract_statisticForm.contractDelBitBtnClick(Sender: TObject);
Var
  contract_id : integer;
  contract_no : string;
begin
  contract_no := allContractGridQuery.fieldbyname('contract_no').asstring;
  if MessageDlg('Вы действительно хотите удалить договор ' + contract_no + ' ?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    contract_id := allContractGridQuery.fieldbyname('contract_id').asinteger;
    with deleteContractQuery do begin
      Close;
      ParamByName('contract_id').asinteger:= contract_id;
      ParamByName('rel_type_id').asfloat := cntr_grp.GetCurRelTypeID;
      ParamByName('ent_type_id').asfloat := cntr_grp.GetCurEntTypeID;
    end;
    deleteContractQuery.ExecSQL;
    // перечитываем данные
    allContractGridQuery.Close;
    allContractGridQuery.Open;
  end;
end;

//---------------------------------------------------------------
// процедура обработки выбора группы договоров
//---------------------------------------------------------------
procedure Tcontract_statisticForm.contract_grpComboBoxChange(
  Sender: TObject);
begin
  // если группа договоров выбрана
  // разрешаем дальнейшую работу
  if ((contract_grpComboBox.ItemIndex <> -1)
      and (contract_grpComboBox.Text <> '')) then begin
    statistic_contractTabSheet.Enabled := true;
    changeTabSheet.Enabled := true;
    readonlyCheckBox.Enabled := true;
    // формируем список всех контрактов выбранного типа
    // для отображения в таблице
    with allContractGridQuery do begin
      Close;
      ParamByName('rel_type_id').asfloat := cntr_grp.GetCurRelTypeID;
      ParamByName('ent_type_id').asfloat := cntr_grp.GetCurEntTypeID;
      Open;
    end;
  end;
end;

//---------------------------------------------------------------
// обрабатываем кнопку ReadOnly
//---------------------------------------------------------------
procedure Tcontract_statisticForm.readonlyCheckBoxClick(Sender: TObject);
begin
  if readonlyCheckBox.Checked then begin
    contractAddBitBtn.Enabled := false;
    contractDelBitBtn.Enabled := false;
  end
  else begin
    contractAddBitBtn.Enabled := true;
    contractDelBitBtn.Enabled := true;
  end;
end;

//---------------------------------------------------------------
// обрабатываем вход в закладку "Группа договоров"
//---------------------------------------------------------------
procedure Tcontract_statisticForm.choose_contractTabSheetEnter(
  Sender: TObject);
begin
  sbReportToExcel.Enabled := false;
end;

//---------------------------------------------------------------
// обрабатываем вход в закладку "Статистика по договорам"
//---------------------------------------------------------------
procedure Tcontract_statisticForm.statistic_contractTabSheetEnter(
  Sender: TObject);
begin
  // если группа договоров выбрана
  // разрешаем сформировать отчет
  if ((contract_grpComboBox.ItemIndex <> -1)
      and (contract_grpComboBox.Text <> '')
      and not (sbReportToExcel.Enabled)) then begin
    sbReportToExcel.Enabled := true;
  end;
end;

//---------------------------------------------------------------
// обрабатываем вход в закладку "Отслеживание изменений"
//---------------------------------------------------------------
procedure Tcontract_statisticForm.changeTabSheetEnter(Sender: TObject);
begin
  // если группа договоров выбрана
  // разрешаем сформировать отчет
  if ((contract_grpComboBox.ItemIndex <> -1)
      and (contract_grpComboBox.Text <> '')
      and not (sbReportToExcel.Enabled)) then begin
    sbReportToExcel.Enabled := true;
  end;
end;

end.
