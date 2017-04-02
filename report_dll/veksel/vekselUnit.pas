unit vekselUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin;

const
  sAllPage = 'forAllTabSheet';
  sEnterprPage = 'forEnterprTabSheet';
  sSaldoPayVekselPage = 'forSaldoPayVekselTabSheet';
  sSaldoSaleVekselPage = 'forSaldoSaleVekselTabSheet';
  sChangeVekselPage = 'changeVekselTabSheet';
  sVekselTemplate = 'veksel.xlt';
  sSaldoVekselTemplate = 'saldo_veksel.xlt';
  sChangeVekselTemplate = 'change.xlt';

type
  TVekselExportForm = class(TForm)
    VekselPageControl: TPageControl;
    allVekselQuery: TQuery;
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
    allVekselInContractQuery: TQuery;
    GetContractDateQuery: TQuery;
    changeVekselTabSheet: TTabSheet;
    JournalDateMaskEdit: TMaskEdit;
    Label3: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    changeVekselQuery: TQuery;
    forSaldoSaleVekselTabSheet: TTabSheet;
    Label10: TLabel;
    allVekselOutContractQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure ExportVeksel(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure ExportVekselInSaldo(Sender: TObject);
    procedure ExportVekselOutSaldo(Sender: TObject);
    procedure ExportChangeVeksel(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure forSaldoPayVekselTabSheetShow(Sender: TObject);
    procedure forSaldoPayVekselTabSheetHide(Sender: TObject);
    procedure forSaldoSaleVekselTabSheetShow(Sender: TObject);
    procedure forSaldoSaleVekselTabSheetHide(Sender: TObject);
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

procedure TVekselExportForm.FormShow(Sender: TObject);
begin
  VekselBeginMaskEdit.Text := startDate;
  VekselEndMaskEdit.Text := DateToStr(Date);
  JournalDateMaskEdit.Text := DateToStr(Date-1);
end;

//---------------------------------------------------------------------
// выбор векселей по предприятию
//---------------------------------------------------------------------
procedure TVekselExportForm.ExportVeksel(Sender: TObject);
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
     countVeksel : integer ;

     creditor_name : string;
     debitor_name : string;
//     act_no : string;
     pay_date : TDate;
     amount : real;
     amount_usd : real;
     veksel_type : string;
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

    PathToTemplate := PathToProgram + '\Template\' + sVekselTemplate;
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
     countVeksel := 0;
     row := 4;

     { просим в базе необходимые счета }
     allVekselQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allVekselQuery.Eof do begin
       countVeksel := countVeksel + 1;

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       // master
       creditor_name := allVekselQuery.fieldbyname('creditor').asstring;
       debitor_name := allVekselQuery.fieldbyname('debitor').asstring;
//       act_no :=;
       pay_date := allVekselQuery.fieldbyname('pay_date').asdatetime;
       amount := allVekselQuery.fieldbyname('amounthrivn').asfloat;
       amount_usd := allVekselQuery.fieldbyname('amount_usd').asfloat;
       veksel_type := allVekselQuery.fieldbyname('type_name').asstring;
       contract_no := allVekselQuery.fieldbyname('contract_no').asstring;
       comment := allVekselQuery.fieldbyname('comments').asstring;

//       s_row := IntToStr(row);
       info_row[1] := countVeksel;
       info_row[2] := creditor_name;
       info_row[3] := debitor_name;
//       info_row[4] := act_no;
       info_row[5] := pay_date;
       info_row[6] := amount;
       info_row[7] := amount_usd;
       info_row[8] := veksel_type;
       info_row[9] := contract_no;
       info_row[10] := comment;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'J' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       for i := 1 to 10 do info_row[i] := unAssigned;

       row := row + 1;
       allVekselQuery.Next;
     end;

   finally
     Excel.free;
     allVekselQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

// -----------------------------------------------------------------------
// формирует отчет по договорам купли продажи векселей
// только договора в которых векселя заходили на нас
// -----------------------------------------------------------------------
procedure TVekselExportForm.ExportVekselInSaldo(Sender: TObject);
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
  count_enterpr : integer;
  cur_enterpr_id : real;
  prev_enterpr_id : real;
  SaldoOfContract : real;

  enterprise_name : string;
  SigningDate : TDateTime;
  contract_veksel_pay : string;
  debit : real;
  credit : real;

const
   English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
   Column = 0;
begin
  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  PathToTemplate := PathToProgram + '\Template\' + sSaldoVekselTemplate;
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
    row := 5;

    count_enterpr := 0;
    prev_enterpr_id := -1;
    with allVekselInContractQuery do
      ParamByName('end_date').asdate := EndDate;

    allVekselInContractQuery.Open;

    while not allVekselInContractQuery.Eof do begin
      cur_enterpr_id := allVekselInContractQuery.fieldbyname('enterprise_id').asfloat;
      enterprise_name := allVekselInContractQuery.fieldbyname('enterprise_name').asString;
      row := row + 1;
      SaldoOfContract := 0;
      if prev_enterpr_id <> cur_enterpr_id then begin
        count_enterpr := count_enterpr + 1;
        info_row[1] := count_enterpr;
        info_row[2] := Enterprise_name;
      end;

      contract_veksel_pay := allVekselInContractQuery.fieldbyname('contract').asString;
      debit := allVekselInContractQuery.fieldbyname('debit').asfloat;
      credit := allVekselInContractQuery.fieldbyname('credit').asfloat;
      SaldoOfContract := debit - credit;

      with GetContractDateQuery do begin
        Close;
        ParamByName('contract_no').asstring := contract_veksel_pay;
      end;
      GetContractDateQuery.Open;
      SigningDate := GetContractDateQuery.fieldbyname('signing_date').asdatetime;

      if contract_veksel_pay <> 'нет контракта' then begin
        info_row[3] := contract_veksel_pay;
        info_row[4] := SigningDate;
        info_row[5] := debit;
        info_row[6] := credit;
        info_row[7] := SaldoOfContract;
      end;
      cellFrom := 'A' + IntToStr(row);
      cellTo := 'G' + IntToStr(row);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 10 do info_row[i] := unAssigned;

      prev_enterpr_id := cur_enterpr_id;
      allVekselInContractQuery.Next;
    end;  // end of while for allVekselInContractQuery

  finally
    Excel.free;
    allVekselInContractQuery.Close;
    GetContractDateQuery.Close;
    SetThreadLocale(Temp);
  end;
end;

// -----------------------------------------------------------------------
// формирует отчет по договорам купли продажи векселей
// только договора в которых векселя уходили от нас
// -----------------------------------------------------------------------
procedure TVekselExportForm.ExportVekselOutSaldo(Sender: TObject);
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
  count_enterpr : integer;
  cur_enterpr_id : real;
  prev_enterpr_id : real;
  SaldoOfContract : real;

  enterprise_name : string;
  SigningDate : TDateTime;
  contract_veksel_pay : string;
  debit : real;
  credit : real;

const
   English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
   Column = 0;
begin
  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  PathToTemplate := PathToProgram + '\Template\' + sSaldoVekselTemplate;
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
    row := 5;

    count_enterpr := 0;
    prev_enterpr_id := -1;
    with allVekselOutContractQuery do
      ParamByName('end_date').asdate := EndDate;

    allVekselOutContractQuery.Open;

    while not allVekselOutContractQuery.Eof do begin
      cur_enterpr_id := allVekselOutContractQuery.fieldbyname('enterprise_id').asfloat;
      enterprise_name := allVekselOutContractQuery.fieldbyname('enterprise_name').asString;
      row := row + 1;
      SaldoOfContract := 0;
      if prev_enterpr_id <> cur_enterpr_id then begin
        count_enterpr := count_enterpr + 1;
        info_row[1] := count_enterpr;
        info_row[2] := Enterprise_name;
      end;

      contract_veksel_pay := allVekselOutContractQuery.fieldbyname('contract').asString;
      debit := allVekselOutContractQuery.fieldbyname('debit').asfloat;
      credit := allVekselOutContractQuery.fieldbyname('credit').asfloat;
      SaldoOfContract := debit - credit;

      with GetContractDateQuery do begin
        Close;
        ParamByName('contract_no').asstring := contract_veksel_pay;
      end;
      GetContractDateQuery.Open;
      SigningDate := GetContractDateQuery.fieldbyname('signing_date').asdatetime;

      if contract_veksel_pay <> 'нет контракта' then begin
        info_row[3] := contract_veksel_pay;
        info_row[4] := SigningDate;
        info_row[5] := debit;
        info_row[6] := credit;
        info_row[7] := SaldoOfContract;
      end;
      cellFrom := 'A' + IntToStr(row);
      cellTo := 'G' + IntToStr(row);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 10 do info_row[i] := unAssigned;

      prev_enterpr_id := cur_enterpr_id;
      allVekselOutContractQuery.Next;
    end;  // end of while for allVekselOutContractQuery

  finally
    Excel.free;
    allVekselOutContractQuery.Close;
    GetContractDateQuery.Close;
    SetThreadLocale(Temp);
  end;
end;

//---------------------------------------------------------------------
// формирование отчета об изменениях в базе за отчетный период
// на определенную дату
//---------------------------------------------------------------------
procedure TVekselExportForm.ExportChangeVeksel(Sender: TObject);
Var
  temp: lcid;
  Excel : TExcel;
  cell : string;
  cellFrom : string;
  cellTo : string;
  info_row : array[1..15] of Variant;
  PathToTemplate : string;
  i : integer;
  row : integer;
  JournalDate : TDate;
  //
  countChange : integer ;
  //
  type_journal : integer;
  type_name_journal : string;
  user_name : string;
  j_pay_date : TDate;
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
  PathToTemplate := PathToProgram + '\Template\' + sChangeVekselTemplate;
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

    with changeVekselQuery do begin
      Close;
      ParamByName('pay_begin_date').asdate := BeginDate;
      ParamByName('pay_end_date').asdate := EndDate;
      ParamByName('journal_date').asdate := JournalDate;
    end;
    changeVekselQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
    while not changeVekselQuery.Eof do begin
      countChange := countChange + 1;

    // ----- ------
      Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
      type_journal := changeVekselQuery.fieldbyname('type').asinteger;
      case type_journal of
        1 : type_name_journal := 'удаление';
        2 : type_name_journal := 'изменение';
        3 : type_name_journal := 'добавление';
      end;
      user_name := changeVekselQuery.fieldbyname('user_name').asstring;
      j_pay_date := changeVekselQuery.fieldbyname('o_pay_date').asdatetime;
      j_amount := changeVekselQuery.fieldbyname('o_summa').asfloat;
      j_contract_no := changeVekselQuery.fieldbyname('o_contract_no').asstring;
      journal_date := changeVekselQuery.fieldbyname('journal_date').asdatetime;
      debitor_name := changeVekselQuery.fieldbyname('debitor').asstring;
      creditor_name := changeVekselQuery.fieldbyname('creditor').asstring;
      type_name := changeVekselQuery.fieldbyname('type_name').asstring;
      amount := changeVekselQuery.fieldbyname('amount').asfloat;
      pay_date := changeVekselQuery.fieldbyname('pay_date').asdatetime;
      contract_no := changeVekselQuery.fieldbyname('contract_no').asstring;
      comment := changeVekselQuery.fieldbyname('comments').asstring;

      info_row[1] := type_name_journal;
      info_row[2] := user_name;
      info_row[3] := j_pay_date;
      info_row[4] := j_amount;
      info_row[5] := j_contract_no;
      info_row[6] := journal_date;
      info_row[7] := ' ';
      info_row[8] := debitor_name;
      info_row[9] := creditor_name;
      info_row[10] := type_name;
      info_row[11] := amount;
      info_row[12] := pay_date;
      info_row[13] := contract_no;
      info_row[14] := comment;

      cellFrom := 'A' + IntToStr(row);
      cellTo := 'N' + IntToStr(row);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 15 do info_row[i] := unAssigned;

      row := row + 1;
      changeVekselQuery.Next;
    end;

  finally
    Excel.free;
    changeVekselQuery.Close;
    SetThreadLocale(Temp);
  end;
end;

//---------------------------------------------------------------
procedure TVekselExportForm.sbReportToExcelClick(Sender: TObject);
Var
  id : integer;
  name : string;
  s : array[0..maxPChar] of Char;
  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
  pname := @s;
  BeginDate := StrToDate(VekselBeginMaskEdit.Text);
  EndDate := StrToDate(VekselEndMaskEdit.Text);

  if VekselPageControl.ActivePage.Name = sAllPage then
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
           SQL.Add('AND ((o.type_id = 4) or (o.type_id = 5) or (o.type_id = 15) or (o.type_id = 20) or (o.type_id = 24))');
           SQL.Add('AND o.pay_date >= :begin_date');
           SQL.Add('AND o.pay_date <= :end_date');
           SQL.Add('ORDER BY o.creditor_id, O.PAY_DATE,O.AMOUNTHRIVN, O.COMMENTS');
           Prepare;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end;
         ReportHeader := 'Все векселя за период с ' +
                  VekselBeginMaskEdit.Text + ' по ' + VekselEndMaskEdit.Text;

         // формируем отчет
         ExportVeksel(Sender);
       end; // конец iAllPage

  if VekselPageControl.ActivePage.Name = sEnterprPage then
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
           SQL.Add('AND ((o.type_id = 4) or (o.type_id = 5) or (o.type_id = 15) or (o.type_id = 20) or (o.type_id = 24))');
           SQL.Add('AND o.pay_date >= :begin_date');
           SQL.Add('AND o.pay_date <= :end_date');
           SQL.Add('AND ((o.creditor_id = :ent_id) or (o.debitor_id = :ent_id))');
           SQL.Add('ORDER BY o.creditor_id, O.PAY_DATE,O.AMOUNTHRIVN, O.COMMENTS');
           Prepare;
           ParamByName('begin_date').asdate := BeginDate;
           ParamByName('end_date').asdate := EndDate;
         end;

         if GetEnterprise(id,pname) = mrOk then begin
           name := string(pname);
           allVekselQuery.ParamByName('ent_id').asinteger := id;
         end
         else
          raise Exception.Create('Предприятие не выбрано');

         ReportHeader := 'Векселя за период с ' +
                VekselBeginMaskEdit.Text + ' по ' + VekselEndMaskEdit.Text +
                  ' ' + '(' + name  + ')';
         // формируем отчет
         ExportVeksel(Sender);
       end; // конец sEnterprPage

  if VekselPageControl.ActivePage.Name = sSaldoPayVekselPage then
       begin
         ReportHeader := 'Сальдо по векселям предъявленным нам к оплате и договорам купли векселей  на '
                         + VekselEndMaskEdit.Text +
                         '    (' + TimeToStr(Time) + ')';
         // формируем отчет
         ExportVekselInSaldo(Sender);
       end; // конец sSaldoPayVekselPage

  if VekselPageControl.ActivePage.Name = sSaldoSaleVekselPage then
       begin
         ReportHeader := 'Сальдо по векселям предъявленным нами к оплате и договорам продажи векселей  на '
                         + VekselEndMaskEdit.Text +
                         '    (' + TimeToStr(Time) + ')';
         // формируем отчет
         ExportVekselOutSaldo(Sender);
       end; // конец sSaldoSaleVekselPage

  if VekselPageControl.ActivePage.Name = sChangeVekselPage then
    begin
      ReportHeader := 'Журнал изменений в базе данных ДИСа за отчетный период с '
                      + VekselBeginMaskEdit.Text
                      + ' по '
                      + VekselEndMaskEdit.Text
                      + ' начиная с '
                      + JournalDateMaskEdit.Text;
         // формируем отчет
//      ExportChangeVeksel(Sender);
    end; // конец sChangeVekselPage

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TVekselExportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;


procedure TVekselExportForm.forSaldoPayVekselTabSheetShow(Sender: TObject);
begin
  VekselBeginMaskEdit.Visible := false;
end;

procedure TVekselExportForm.forSaldoPayVekselTabSheetHide(Sender: TObject);
begin
  VekselBeginMaskEdit.Visible := true;
end;

procedure TVekselExportForm.forSaldoSaleVekselTabSheetShow(
  Sender: TObject);
begin
  VekselBeginMaskEdit.Visible := false;
end;

procedure TVekselExportForm.forSaldoSaleVekselTabSheetHide(
  Sender: TObject);
begin
  VekselBeginMaskEdit.Visible := true;
end;

end.
