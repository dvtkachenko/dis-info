unit ndsUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, shared_type, excel_type;

const
  indsGeneralPage = 0;
  iProtocolZchtPage = 1;
  sndsReportTemplate = 'nds_report.xlt';
  sprotocol_zchtReportTemplate = 'protocol_zcht.xlt';
  iMaxDept = 10;

type
  TDept_nds = record
    dept_id : integer;
    dept_name : string;
    amount : real;
    nds : real;
    formulaRow : integer;
  end;
  //
  TNds = record
    _in : array[1..iMaxDept] of TDept_nds;
    _out : array[1..iMaxDept] of TDept_nds;
    countIn, countOut : integer;
    servAmountIn : real;
    servAmountOut : real;
    formulaRowServIn : integer;
    formulaRowServOut : integer;
    payCoal : real;

    // данный НДС будет считаться специальными запросами
    // и должен равняться сумме НДС по всем отделам
    allNdsIn : real;
    allNdsOut : real;
    // данная сумма будет считаться специальными запросами
    // и должна равняться сумме счетов по всем отделам
    allAmountIn : real;
    allAmountOut : real;
  end;
  //
  TndsReportForm = class(TForm)
    InvPageControl: TPageControl;
    allInvQuery: TQuery;
    ndsBeginMaskEdit: TMaskEdit;
    ndsEndMaskEdit: TMaskEdit;
    forndsGeneralTabSheet: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    acceptCheckBox: TCheckBox;
    allDeptInQuery: TQuery;
    allDeptOutQuery: TQuery;
    TestInQuery: TQuery;
    TestOutQuery: TQuery;
    allServInvQuery: TQuery;
    payCoalQuery: TQuery;
    protocol_ndsTabSheet: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    allProtocolZchtQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure ndsReport(Sender: TObject);
    procedure ndsReportToExcel(Excel : TExcel; dept_nds : TDept_nds);
    procedure payCoalReportToExcel(Excel : TExcel; nds : TNds);
    procedure ServiceInvReportToExcel(Excel : TExcel; Var allAmount : real; Var formulaRow : integer);
    procedure mainReportToExcel(Excel : TExcel; nds : TNds);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure ExportProtocolZcht(Sender: TObject);
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

{$R *.DFM}

{сервисные процедуры}

{-------------------}

procedure TndsReportForm.FormShow(Sender: TObject);
begin
  ndsBeginMaskEdit.Text := startDate;
  ndsEndMaskEdit.Text := DateToStr(Date);
end;

//---------------------------------------------------------------------
// головная процедура формирования отчета по НДС
//---------------------------------------------------------------------
procedure TndsReportForm.ndsReport(Sender: TObject);
Var
  temp: lcid;
  Excel : TExcel;
  PathToTemplate : string;
//  i : integer;
  dept_id : integer;
  dept_name : string;
  countDeptIn  : integer;
  countDeptOut : integer;
  nds : TNds;

const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  Column = 0;
begin
  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  PathToTemplate := PathToProgram + '\Template\' + sndsReportTemplate;
  try
    Excel.AddWorkBook(PathToTemplate);
    Excel.Visible := true;
  except
    raise Exception.Create('Невозможно загрузить Excel');
  end;

  try
    countDeptIn  := 0;
    countDeptOut := 0;

    nds.countIn := 0;
    nds.countOut := 0;

    //  вытаскиваем все отделы по которым были входящие
    //  счета-фактуры в заданном периоде
    with allDeptInQuery do begin
      Close;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    allDeptInQuery.Open;

    while not allDeptInQuery.Eof do begin
      countDeptIn := countDeptIn + 1;
      dept_id := allDeptInQuery.fieldbyname('dept_id').asinteger;
      dept_name := allDeptInQuery.fieldbyname('dept_name').asstring;
      ReportHeader := 'Входящие счета-фактуры (' + dept_name + ')'
                      + ' за период с ' + ndsBeginMaskEdit.Text + ' по '
                      + ndsEndMaskEdit.Text;
      with allInvQuery do begin
        Close;
        SQL.Clear;
        SQL.Add('select sender_name enterpr_name,');
        SQL.Add('pay_date,');
        SQL.Add('invoice_date,');
        SQL.Add('invoice_no,');
        SQL.Add('short_trade_mark,');
        SQL.Add('amount,');
        SQL.Add('nds,');
        SQL.Add('contract,');
        SQL.Add('dept_name');
        SQL.Add('from balans_report_all_invoices(:begin_date, :end_date)');
        SQL.Add('where dept_id = :id and payer_id = 0');
        SQL.Add('and is_in_oper = ''Y''');
        SQL.Add('order by sender_name, pay_date');
        Prepare;
        ParamByName('begin_date').asdate := BeginDate;
        ParamByName('end_date').asdate := EndDate;
        ParamByName('id').asinteger := dept_id;
      end;
      allInvQuery.Open;
      // ------------------
      // создание нового листа в Excel
      // это необходимо , чтобы процедура ndsReportToExcel
      // не замечала подмены листа и при каждом последующем вызове
      // формировала отчет на новом листе
      Excel.CopyWorkSheet('source','in_'+ IntToStr(countDeptIn));
      // ------------------

      // вызываем процедуру формирования отчета по отделу в Excel
      nds.countIn := nds.countIn + 1;
      nds._in[nds.countIn].dept_id := dept_id;
      nds._in[nds.countIn].dept_name := dept_name;
      ndsReportToExcel(Excel,nds._in[nds.countIn]);

      allDeptInQuery.Next;
    end; // КОНЕЦ allDeptInQuery

    //  вытаскиваем все отделы по которым были исходящие
    //  счета-фактуры в заданном периоде
    with allDeptOutQuery do begin
      Close;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    allDeptOutQuery.Open;

    while not allDeptOutQuery.Eof do begin
      countDeptOut := countDeptOut + 1;
      dept_id := allDeptOutQuery.fieldbyname('dept_id').asinteger;
      dept_name := allDeptOutQuery.fieldbyname('dept_name').asstring;
      ReportHeader := 'Исходящие счета-фактуры (' + dept_name + ')'
                      + ' за период с ' + ndsBeginMaskEdit.Text + ' по '
                      + ndsEndMaskEdit.Text;
      with allInvQuery do begin
        Close;
        SQL.Clear;
        SQL.Add('select payer_name enterpr_name,');
        SQL.Add('pay_date,');
        SQL.Add('invoice_date,');
        SQL.Add('invoice_no,');
        SQL.Add('short_trade_mark,');
        SQL.Add('amount,');
        SQL.Add('nds,');
        SQL.Add('contract,');
        SQL.Add('dept_name');
        SQL.Add('from balans_report_all_invoices(:begin_date, :end_date)');
        SQL.Add('where dept_id = :id and sender_id = 0');
        SQL.Add('and is_in_oper = ''Y''');
        SQL.Add('order by payer_name, pay_date');
        Prepare;
        ParamByName('begin_date').asdate := BeginDate;
        ParamByName('end_date').asdate := EndDate;
        ParamByName('id').asinteger := dept_id;
      end;
      allInvQuery.Open;
      // ------------------
      // создание нового листа в Excel
      // это необходимо , чтобы процедура ndsReportToExcel
      // не замечала подмены листа и при каждом последующем вызове
      // формировала отчет на новом листе
      // ------------------
      Excel.CopyWorkSheet('source','out_'+ IntToStr(countDeptOut));

      // вызываем процедуру формирования отчета по отделу в Excel
      nds.countOut := nds.countOut + 1;
      nds._in[nds.countOut].dept_id := dept_id;
      nds._in[nds.countOut].dept_name := dept_name;
      ndsReportToExcel(Excel,nds._in[nds.countOut]);

      allDeptOutQuery.Next;
    end; // КОНЕЦ allDeptOutQuery

    //  формирование отчета об оплате углей
    Excel.SelectWorkSheet('оплата углей');
    with PayCoalQuery do begin
      Close;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    PayCoalQuery.Open;
    ReportHeader := ' за период с ' + ndsBeginMaskEdit.Text + ' по '
                      + ndsEndMaskEdit.Text;
    payCoalReportToExcel(Excel,nds);

    // ------------------
    // ------------------
    //  формирование отчета по входящим счетам за услуги
    Excel.SelectWorkSheet('услуги_in');
    with allServInvQuery do begin
      Close;
      SQL.Clear;
      SQL.Add('select e.enterprise_name enterprise_name,');
      SQL.Add('o.pay_date, st.type_name,');
      SQL.Add('o.debit_hrv amount, o.contract_no, o.comments');
      SQL.Add('from operation_list2 o,source_types st, enterpr e');
      SQL.Add('where o.enterpr_id = e.enterpr_id');
      SQL.Add('and type_id = 18');
      SQL.Add('and o.type_id = st.type_id');
      SQL.Add('and o.pay_date >= :begin_date');
      SQL.Add('and o.pay_date <= :end_date');
      SQL.Add('and o.debitor_id = 0');
      SQL.Add('order by e.enterprise_name');
      Prepare;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    allServInvQuery.Open;
    ReportHeader := ' за период с ' + ndsBeginMaskEdit.Text + ' по '
                      + ndsEndMaskEdit.Text;
    ServiceInvReportToExcel(Excel,nds.servAmountIn,nds.formulaRowServIn);

    //  формирование отчета по исходящим счетам за услуги
    Excel.SelectWorkSheet('услуги_out');
    with allServInvQuery do begin
      Close;
      SQL.Clear;
      SQL.Add('select e.enterprise_name enterprise_name,');
      SQL.Add('o.pay_date, st.type_name,');
      SQL.Add('o.debit_hrv amount, o.contract_no, o.comments');
      SQL.Add('from operation_list2 o,source_types st, enterpr e');
      SQL.Add('where o.enterpr_id = e.enterpr_id');
      SQL.Add('and type_id = 18');
      SQL.Add('and o.type_id = st.type_id');
      SQL.Add('and o.pay_date >= :begin_date');
      SQL.Add('and o.pay_date <= :end_date');
      SQL.Add('and o.creditor_id = 0');
      SQL.Add('order by e.enterprise_name');
      Prepare;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    allServInvQuery.Open;
    ReportHeader := ' за период с ' + ndsBeginMaskEdit.Text + ' по '
                      + ndsEndMaskEdit.Text;
    ServiceInvReportToExcel(Excel,nds.servAmountOut,nds.formulaRowServOut);
    // ------------------
    // ------------------

    // ------------------
    // код для расчета контрольных переменных
    // переключаемся на сводный лист отчета по НДС
    Excel.SelectWorkSheet('НДС');
    with TestInQuery do begin
      Close;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    TestInQuery.Open;
    //
    with TestOutQuery do begin
      Close;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
    end;
    TestOutQuery.Open;

    nds.allNdsIn := TestInQuery.fieldbyname('nds').asfloat;
    nds.allAmountIn := TestInQuery.fieldbyname('amount').asfloat;
    nds.allNdsOut :=  TestOutQuery.fieldbyname('nds').asfloat;
    nds.allAmountOut := TestOutQuery.fieldbyname('amount').asfloat;
    mainReportToExcel(Excel,nds);

    // ------------------

  finally
    Excel.free;
    allDeptInQuery.Close;
    allDeptOutQuery.Close;
    TestInQuery.Close;
    TestOutQuery.Close;
    allInvQuery.Close;
    allServInvQuery.Close;
    SetThreadLocale(Temp);
  end;
end;

//---------------------------------------------------------------------
// процедура вывода отчета по НДС в Excel
//---------------------------------------------------------------------
procedure TndsReportForm.ndsReportToExcel(Excel : TExcel;dept_nds : TDept_nds);
  Var
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..10] of Variant;
     i : integer;
     row : integer;
     formulaRow : integer;
     formulaAmount : string;

     { контрольные переменные }
     countInvoices : integer ;
     allDeptAmount : real;
     allDeptNds : real;

     // invoice master
//     invoice_id : integer;
     enterpr_name : string;
     pay_date : TDate;
     invoice_date : TDate;
     invoice_no : string;
     trade_mark : string;
     amount : real;
     nds : real;
     contract : string;
     dept_name : string;
  begin

   try
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     { инициализируем  контрольные переменные }
     countInvoices := 0;
     row := 5;
     allDeptAmount := 0;
     allDeptNds := 0;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allInvQuery.Eof do begin
       countInvoices := countInvoices + 1;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

//       invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
       enterpr_name := allInvQuery.fieldbyname('enterpr_name').asstring;
       pay_date := allInvQuery.fieldbyname('pay_date').asdatetime;
       invoice_date := allInvQuery.fieldbyname('invoice_date').asdatetime;
       invoice_no := allInvQuery.fieldbyname('invoice_no').asstring;
       invoice_no := ' ' + invoice_no;
       trade_mark := allInvQuery.fieldbyname('short_trade_mark').asstring;
       amount := allInvQuery.fieldbyname('amount').asfloat;
       nds := allInvQuery.fieldbyname('nds').asfloat;
       contract := allInvQuery.fieldbyname('contract').asstring;
       dept_name := allInvQuery.fieldbyname('dept_name').asstring;

       allDeptAmount := allDeptAmount + amount;
       allDeptNds := allDeptNds + nds;

       info_row[1] := enterpr_name;
       info_row[2] := pay_date;
       info_row[3] := invoice_date;
       info_row[4] := invoice_no;
       info_row[5] := trade_mark;
       info_row[6] := amount;
       info_row[7] := nds;
       info_row[8] := contract;
       info_row[9] := dept_name;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'I' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 10 do info_row[i] := unAssigned;

       row := row + 1;
       allInvQuery.Next;
     end;
     dept_nds.amount := allDeptAmount;
     dept_nds.nds := allDeptnds;

     formulaRow := row + 3;
     dept_nds.formulaRow := formulaRow;
     formulaAmount := '=SUM(R[-' + IntToStr(formulaRow - 5) + ']C:R[-1]C)';
     cell := 'F' + IntToStr(formulaRow);
     Excel.CellFormulaR1C1[cell] := formulaAmount;
     cell := 'G' + IntToStr(formulaRow);
     Excel.CellFormulaR1C1[cell] := formulaAmount;

   finally
     allInvQuery.Close;
   end;
end;

//---------------------------------------------------------------------
// формирование отчета оплаты поставленных углей по договорам
//---------------------------------------------------------------------
procedure TndsReportForm.payCoalReportToExcel(Excel : TExcel; nds : TNds);
  Var
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..10] of Variant;
     i : integer;
     row : integer;

     { контрольные переменные }
     countPay : integer ;
     allPayAmount : real;

     // invoice master
//     invoice_id : integer;
     enterprise_name : string;
     type_name : string;
     amount : real;
     pay_date : TDate;
     contract : string;
     comments : string;
  begin

   try
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     { инициализируем  контрольные переменные }
     countPay := 0;
     row := 5;
     allPayAmount := 0;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not PayCoalQuery.Eof do begin
       countPay := countPay + 1;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

//       invoice_id := allInvQuery.fieldbyname('invoice_id').asinteger;
       enterprise_name := payCoalQuery.fieldbyname('enterprise_name').asstring;
       type_name := payCoalQuery.fieldbyname('type_name').asstring;
       amount := payCoalQuery.fieldbyname('amount').asfloat;
       pay_date := payCoalQuery.fieldbyname('pay_date').asdatetime;
       contract := payCoalQuery.fieldbyname('contract_no').asstring;
       comments := payCoalQuery.fieldbyname('comments').asstring;

       allPayAmount := allPayAmount + amount;

       info_row[1] := enterprise_name;
       info_row[2] := ' ';
       info_row[3] := type_name;
       info_row[4] := amount;
       info_row[5] := pay_date;
       info_row[6] := contract;
       info_row[7] := comments;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'G' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 10 do info_row[i] := unAssigned;

       row := row + 1;
       PayCoalQuery.Next;
     end;
     nds.payCoal := allPayAmount;

   finally
     PayCoalQuery.Close;
   end;
end;

//---------------------------------------------------------------------
// формирование отчета о счетах за услуги
//---------------------------------------------------------------------
procedure TndsReportForm.ServiceInvReportToExcel(Excel : TExcel; var allAmount : real; Var formulaRow : integer);
  Var
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..10] of Variant;
     i : integer;
     row : integer;
     formulaAmount : string;

     { контрольные переменные }
     countServInv : integer ;
     allServInvAmount : real;

     // invoice master
//     invoice_id : integer;
     enterprise_name : string;
     pay_date : TDate;
     type_name : string;
     amount : real;
     contract : string;
     comments : string;
  begin

   try
     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     { инициализируем  контрольные переменные }
     countServInv := 0;
     row := 5;
     allServInvAmount := 0;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allServInvQuery.Eof do begin
       countServInv := countServInv + 1;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

       enterprise_name := allServInvQuery.fieldbyname('enterprise_name').asstring;
       pay_date := allServInvQuery.fieldbyname('pay_date').asdatetime;
       type_name := allServInvQuery.fieldbyname('type_name').asstring;
       amount := allServInvQuery.fieldbyname('amount').asfloat;
       contract := allServInvQuery.fieldbyname('contract_no').asstring;
       comments := allServInvQuery.fieldbyname('comments').asstring;

       allServInvAmount := allServInvAmount + amount;

       info_row[1] := enterprise_name;
       info_row[2] := pay_date;
       info_row[3] := type_name;
       info_row[4] := amount;
       info_row[5] := unAssigned;
       info_row[6] := contract;
       info_row[7] := comments;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'G' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 10 do info_row[i] := unAssigned;

       row := row + 1;
       allServInvQuery.Next;
     end;
     allAmount := allServInvAmount;
     formulaRow := row + 3;
     formulaAmount := '=SUM(R[-' + IntToStr(formulaRow - 5) + ']C:R[-1]C)';
     cell := 'D' + IntToStr(formulaRow);
     Excel.CellFormulaR1C1[cell] := formulaAmount;

   finally
     allServInvQuery.Close;
   end;
end;

//---------------------------------------------------------------------
// формирование сводного отчета по НДС
//---------------------------------------------------------------------
procedure TndsReportForm.mainReportToExcel(Excel : TExcel; nds : TNds);
  Var
     info_row : Variant;
  begin

   try
     info_row := VarArrayCreate([1,2,1,3],varVariant);
     info_row[1,1] := nds.allAmountIn;
     info_row[1,3] := nds.allAmountOut;
     info_row[2,1] := nds.allNdsIn;
     info_row[2,3] := nds.allNdsOut;

     Excel.Range[OleVariant('test'),EmptyParam] := info_row;

   finally
   end;
end;

//---------------------------------------------------------------------
// выбор всех протоколов бюджетного зачета
//---------------------------------------------------------------------
procedure TndsReportForm.ExportProtocolZcht(Sender: TObject);
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

     { контрольные переменные }
     countProtocol : integer ;

     creditor_name : string;
     debitor_name : string;
//     act_no : string;
     pay_date : TDate;
     amount : real;
     amount_usd : real;
     protocol_type : string;
     contract_no : string;
     comment : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    try
      Excel := TExcel.Create;
    except
      raise Exception.Create('Невозможно создать OLE - объект');
    end;

    PathToTemplate := PathToProgram + '\Template\' + sprotocol_zchtReportTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try
     row := 1;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     { инициализируем  контрольные переменные }
     countProtocol := 0;
     row := 4;

     allProtocolZchtQuery.ParamByName('begin_date').asdate := BeginDate;
     allProtocolZchtQuery.ParamByName('end_date').asdate := EndDate;
     { просим в базе необходимые счета }
     allProtocolZchtQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allProtocolZchtQuery.Eof do begin
       countProtocol := countProtocol + 1;

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       // master
       creditor_name := allProtocolZchtQuery.fieldbyname('creditor').asstring;
       debitor_name := allProtocolZchtQuery.fieldbyname('debitor').asstring;
//       act_no :=;
       pay_date := allProtocolZchtQuery.fieldbyname('pay_date').asdatetime;
       amount := allProtocolZchtQuery.fieldbyname('amounthrivn').asfloat;
       amount_usd := allProtocolZchtQuery.fieldbyname('amount_usd').asfloat;
       protocol_type := allProtocolZchtQuery.fieldbyname('type_name').asstring;
       contract_no := allProtocolZchtQuery.fieldbyname('contract_no').asstring;
       comment := allProtocolZchtQuery.fieldbyname('comments').asstring;

//       s_row := IntToStr(row);
       info_row[1] := countProtocol;
       info_row[2] := creditor_name;
       info_row[3] := debitor_name;
//       info_row[4] := act_no;
       info_row[5] := pay_date;
       info_row[6] := amount;
       info_row[7] := amount_usd;
       info_row[8] := protocol_type;
       info_row[9] := contract_no;
       info_row[10] := comment;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'J' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       for i := 1 to 10 do info_row[i] := unAssigned;

       row := row + 1;
       allProtocolZchtQuery.Next;
     end;

   finally
     Excel.free;
     allProtocolZchtQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

// обработчик нажатия кнопки на панели инструментов
procedure TndsReportForm.sbReportToExcelClick(Sender: TObject);
begin
  { конструирование запросов }
  BeginDate := StrToDate(ndsBeginMaskEdit.Text);
  EndDate := StrToDate(ndsEndMaskEdit.Text);

  case InvPageControl.ActivePage.TabIndex of

    indsGeneralPage :
       begin
         ndsReport(Sender);
       end; // конец indsPage

    iProtocolZchtPage :
       begin
         ReportHeader := 'Все бюджетные зачеты по ДИСу за период с ' +
                  ndsBeginMaskEdit.Text + ' по ' + ndsEndMaskEdit.Text ; 
         ExportProtocolZcht(Sender);
       end; // конец iProtocolZchtPage
  end;  // end of CASE

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TndsReportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
