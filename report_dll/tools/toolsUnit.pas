unit toolsUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, shared_type;

const
  icontract_relPage = 0;
  ino_contract_relPage = 1;
  savg_ratePage = 'avg_rateTabSheet';
  sChangeOperationsPage = 'changeOperationsTabSheet';
  savg_rateTemplate = 'avg_rate.xlt';
  sChangeOperationsTemplate = 'change.xlt';

type
  TtoolsForm = class(TForm)
    toolsPageControl: TPageControl;
    avg_rate_disQuery: TQuery;
    avg_rateTabSheet: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    ruleGroupBox: TGroupBox;
    avg_rate_isdQuery: TQuery;
    get_changeOperationsQuery: TQuery;
    changeOperationsTabSheet: TTabSheet;
    arBeginMaskEdit: TMaskEdit;
    arEndMaskEdit: TMaskEdit;
    chopBeginMaskEdit: TMaskEdit;
    chopEndMaskEdit: TMaskEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    JournalDateMaskEdit: TMaskEdit;
    currencyRadioGroup: TRadioGroup;
    procedure FormShow(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure avg_rateReport(Sender: TObject);
    procedure ExportChangeOperations(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    parentConfig : p_config;
    ReportHeader : string;
    BeginDate : TDateTime;
    EndDate : TDateTime;
    PathToProgram : string;
  end;

implementation

uses excel_type;

{$R *.DFM}

{сервисные процедуры}

{-------------------}

procedure TtoolsForm.FormShow(Sender: TObject);
begin
  arBeginMaskEdit.Text := startDate;
  arEndMaskEdit.Text := DateToStr(Date);
end;

//---------------------------------------------------------------------
// вытаскивает курс доллара за указываемый период и вычисляет его среднее значение
//---------------------------------------------------------------------
procedure TtoolsForm.avg_rateReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..3] of Variant;
     PathToTemplate : string;
     i : integer;
//     ReportHeader : string;
     row : integer;

     { контрольные переменные }
     count_rate: integer ;
     sum_rate : real;
     //
     currency_text : string;
     currency_dis_id : integer;
     currency_isd_id : integer;
     rate_date : TDate;
     rate : real;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + savg_rateTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try

     currency_dis_id := 1;
     currency_isd_id := 840;
     currency_text := 'доллара';

     // выбран доллар
     if currencyRadioGroup.ItemIndex = 0 then begin
       currency_text := 'доллара';
       currency_dis_id := 1;
       currency_isd_id := 840;
     end;

     // выбран евро
     if currencyRadioGroup.ItemIndex = 1 then begin
       currency_text := 'евро';
       currency_dis_id := 978;
       currency_isd_id := 978;
     end;

     // выбран рубль
     if currencyRadioGroup.ItemIndex = 2 then begin
       currency_text := 'рубля';
       currency_dis_id := 810;
       currency_isd_id := 810;
     end;

     ReportHeader := 'Курс ' + currency_text + ' за период с '
                     + arBeginMaskEdit.Text +
                     ' по ' +  arEndMaskEdit.Text;

     row := 2;
     cell := 'A' + IntToStr(row);
     Excel.Cell[cell] := ReportHeader;

     { инициализируем  контрольные переменные
       для вытяжки курса из БД ДИС 98 }
     count_rate := 0;
     sum_rate := 0;
     row := 6;

     try
       { просим в базе ДИС 98 необходимые данные }
       with avg_rate_disQuery do begin
         Close;
         Prepare;
         ParamByName('cur_id').asinteger := currency_dis_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
       end;
       avg_rate_disQuery.Open;

    // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
       while not avg_rate_disQuery.Eof do begin
         count_rate := count_rate + 1;
      // ----- ------
         Update;
      // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
         rate_date := avg_rate_disQuery.fieldbyname('rate_date').asdatetime;
         rate := avg_rate_disQuery.fieldbyname('rate').asfloat;

         sum_rate := sum_rate + rate;

         info_row[1] := count_rate;
         info_row[2] := rate_date;
         info_row[3] := rate;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'C' + IntToStr(row);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 3 do info_row[i] := unAssigned;

         row := row + 1;
         avg_rate_disQuery.Next;
       end;

       row := row + 2;

       for i := 1 to 3 do info_row[i] := unAssigned;
       cellFrom := 'A' + IntToStr(row);
       cellTo := 'C' + IntToStr(row);

       info_row[2] := 'Средний курс (БД ДИС):';
       info_row[3] := sum_rate/count_rate;

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       //
     except
       MessageDlg('Не удается извлечь курс из базы ДИС 98', mtError, [mbOk], 0);
     end;

     Application.BringToFront;
     try
       { просим в базе ИСД 2000 необходимые данные }
       with avg_rate_isdQuery do begin
         Close;
         Prepare;
         ParamByName('cur_id').asinteger := currency_isd_id;
         ParamByName('begin_date').asdate := BeginDate;
         ParamByName('end_date').asdate := EndDate;
       end;
       avg_rate_isdQuery.Open;

       { инициализируем  контрольные переменные
         для вытяжки курса из ИСД 2000 }
       count_rate := 0;
       sum_rate := 0;
       row := 6;

    // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
       while not avg_rate_isdQuery.Eof do begin
         count_rate := count_rate + 1;
      // ----- ------
         Update;
      // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
         rate_date := avg_rate_isdQuery.fieldbyname('rate_date').asdatetime;
         rate := avg_rate_isdQuery.fieldbyname('rate').asfloat;

         sum_rate := sum_rate + rate;

         info_row[1] := rate_date;
         info_row[2] := rate;

         cellFrom := 'D' + IntToStr(row);
         cellTo := 'E' + IntToStr(row);

         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

         for i := 1 to 3 do info_row[i] := unAssigned;

         row := row + 1;
         avg_rate_isdQuery.Next;
       end;

       row := row + 2;

       for i := 1 to 3 do info_row[i] := unAssigned;
       cellFrom := 'D' + IntToStr(row);
       cellTo := 'E' + IntToStr(row);

       info_row[1] := 'Средний курс (ИСД 2000):';
       info_row[2] := sum_rate/count_rate;

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

     except
       MessageDlg('Не удается извлечь курс из базы ИСД 2000', mtError, [mbOk], 0);
     end;

   finally
     Excel.free;
     avg_rate_disQuery.Close;
     avg_rate_isdQuery.Close;
     SetThreadLocale(Temp);
   end;
end;

//---------------------------------------------------------------------
// формирование отчета об изменениях в базе за отчетный период
// на определенную дату
//---------------------------------------------------------------------
procedure TtoolsForm.ExportChangeOperations(Sender: TObject);
Var
  temp: lcid;
  Excel : TExcel;
  cell : string;
  cellFrom : string;
  cellTo : string;
  info_row : array[1..23] of Variant;
  PathToTemplate : string;
  i : integer;
  row : integer;
  JournalDate : TDate;
  //
  countChange : integer ;
  //
  // информация об изменении
  oj_operation_id : integer;
  o_source_id : integer;
  oj_type_change : integer;
  oj_type_name_change : string;
  is_last_change : string;
  is_mean_change : string;
  oj_user_name : string;
  oj_journal_date : TDate;

  // информация ДО изменения
  oj_debitor : string;
  oj_creditor : string;
  oj_type_name : string;
  oj_pay_date : TDate;
  oj_amount : real;
  oj_contract_no : string;

  // информация ПОСЛЕ изменения
  o_debitor : string;
  o_creditor : string;
  o_type_name : string;
  o_pay_date : TDate;
  o_amount : real;
  o_contract_no : string;

  // дополнительная информация
  o_comment : string;
  o_short_trade_mark : string;
  o_department : string;

const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  // положение колонок в отчете
  oj_debitor_col = 'H';
  oj_creditor_col = 'I';
  oj_type_name_col = 'J';
  oj_pay_date_col = 'K';
  oj_amount_col = 'L';
  oj_contract_no_col = 'M';
  o_debitor_col = 'O';
  o_creditor_col = 'P';
  o_type_name_col = 'Q';
  o_pay_date_col = 'R';
  o_amount_col = 'S';
  o_contract_no_col = 'T';

begin
  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  PathToTemplate := PathToProgram + '\Template\' + sChangeOperationsTemplate;
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

    row := 7;
    JournalDate := StrToDate(JournalDateMaskEdit.Text);

    with get_changeOperationsQuery do begin
      Close;
      ParamByName('begin_date').asdate := BeginDate;
      ParamByName('end_date').asdate := EndDate;
      ParamByName('journal_date').asdate := JournalDate;
    end;
    get_changeOperationsQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
    while not get_changeOperationsQuery.Eof do begin
      countChange := countChange + 1;

    // ----- ------
      Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //

     // информация об изменении
      oj_operation_id := get_changeOperationsQuery.fieldbyname('oj_operation_id').asinteger;
      o_source_id := get_changeOperationsQuery.fieldbyname('o_source_id').asinteger;
      oj_type_change := get_changeOperationsQuery.fieldbyname('oj_type_operation').asinteger;
      case oj_type_change of
        1 : oj_type_name_change := 'удаление';
        2 : oj_type_name_change := 'изменение';
        3 : oj_type_name_change := 'добавление';
      end;
      is_last_change := get_changeOperationsQuery.fieldbyname('is_last_change').asstring;
      is_mean_change := get_changeOperationsQuery.fieldbyname('is_mean_change').asstring;
      oj_user_name := get_changeOperationsQuery.fieldbyname('oj_user_name').asstring;
      oj_journal_date := get_changeOperationsQuery.fieldbyname('oj_journal_date').asdatetime;
      // информация ДО изменения
      oj_debitor := get_changeOperationsQuery.fieldbyname('oj_debitor').asstring;
      oj_creditor := get_changeOperationsQuery.fieldbyname('oj_creditor').asstring;
      oj_type_name := get_changeOperationsQuery.fieldbyname('oj_type_name').asstring;
      oj_pay_date := get_changeOperationsQuery.fieldbyname('oj_pay_date').asdatetime;
      oj_amount := get_changeOperationsQuery.fieldbyname('oj_amount').asfloat;
      oj_contract_no := get_changeOperationsQuery.fieldbyname('oj_contract_no').asstring;
      // информация ПОСЛЕ изменения
      o_debitor := get_changeOperationsQuery.fieldbyname('o_debitor').asstring;
      o_creditor := get_changeOperationsQuery.fieldbyname('o_creditor').asstring;
      o_type_name := get_changeOperationsQuery.fieldbyname('o_type_name').asstring;
      o_pay_date := get_changeOperationsQuery.fieldbyname('o_pay_date').asdatetime;
      o_amount := get_changeOperationsQuery.fieldbyname('o_amount').asfloat;
      o_contract_no := get_changeOperationsQuery.fieldbyname('o_contract_no').asstring;
      // дополнительная информация
      o_comment := get_changeOperationsQuery.fieldbyname('o_comments').asstring;
      o_short_trade_mark := get_changeOperationsQuery.fieldbyname('o_short_trade_mark').asstring;
      o_department := get_changeOperationsQuery.fieldbyname('o_department').asstring;

      // информация об изменении
      info_row[1] := oj_operation_id;
      info_row[2] := o_source_id;
      info_row[3] := oj_type_name_change;
      info_row[4] := is_last_change;
      info_row[5] := is_mean_change;
      info_row[6] := oj_user_name;
      info_row[7] := oj_journal_date;
      // информация ДО изменения
      info_row[8] := oj_debitor;
      info_row[9] := oj_creditor;
      info_row[10] := oj_type_name;
      info_row[11] := oj_pay_date;
      info_row[12] := oj_amount;
      info_row[13] := oj_contract_no;

      info_row[14] := ' ';

      // информация ПОСЛЕ изменения
      info_row[15] := o_debitor;
      info_row[16] := o_creditor;
      info_row[17] := o_type_name;
      info_row[18] := o_pay_date;
      info_row[19] := o_amount;
      info_row[20] := o_contract_no;
      // дополнительная информация
      info_row[21] := o_comment;
      info_row[22] := o_short_trade_mark;
      info_row[23] := o_department;

      cellFrom := 'A' + IntToStr(row);
      cellTo := 'W' + IntToStr(row);

      Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
      for i := 1 to 23 do info_row[i] := unAssigned;

      // пометка записи если это было последнее изменение
      if (is_last_change = 'Y') then begin
        cellFrom := o_debitor_col + IntToStr(row);
        cellTo := o_contract_no_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 35);
      end;

      // нахождение изменений и выделение их желтым цветом
      // в дебиторах
      if (oj_debitor <> o_debitor) then begin
        cellFrom := oj_debitor_col + IntToStr(row);
        cellTo := oj_debitor_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
        cellFrom := o_debitor_col + IntToStr(row);
        cellTo := o_debitor_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
      end;
      // в кредиторах
      if (oj_creditor <> o_creditor) then begin
        cellFrom := oj_creditor_col + IntToStr(row);
        cellTo := oj_creditor_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
        cellFrom := o_creditor_col + IntToStr(row);
        cellTo := o_creditor_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
      end;
      //  в типе операции
      if (oj_type_name <> o_type_name) then begin
        cellFrom := oj_type_name_col + IntToStr(row);
        cellTo := oj_type_name_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
        cellFrom := o_type_name_col + IntToStr(row);
        cellTo := o_type_name_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
      end;
      //  в учетной дате
      if (oj_pay_date <> o_pay_date) then begin
        cellFrom := oj_pay_date_col + IntToStr(row);
        cellTo := oj_pay_date_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
        cellFrom := o_pay_date_col + IntToStr(row);
        cellTo := o_pay_date_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
      end;
      //  в сумме
      if (Round(oj_amount*100) <> Round(o_amount*100)) then begin
        cellFrom := oj_amount_col + IntToStr(row);
        cellTo := oj_amount_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
        cellFrom := o_amount_col + IntToStr(row);
        cellTo := o_amount_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
      end;
      //  в договоре
      if (oj_contract_no <> o_contract_no) then begin
        cellFrom := oj_contract_no_col + IntToStr(row);
        cellTo := oj_contract_no_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
        cellFrom := o_contract_no_col + IntToStr(row);
        cellTo := o_contract_no_col + IntToStr(row);
        Excel.FillRangeColor(cellFrom, cellTo, 6);
      end;

      row := row + 1;
      get_changeOperationsQuery.Next;
    end;

  finally
    Excel.free;
    get_changeOperationsQuery.Close;
    SetThreadLocale(Temp);
  end;
end;

procedure TtoolsForm.sbReportToExcelClick(Sender: TObject);
//Var
//  id : integer;
//  name : string;
//  s : array[0..maxPChar] of Char;
//  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
//  pname := @s;

  if toolsPageControl.ActivePage.Name = savg_ratePage then
    begin
      BeginDate := StrToDate(arBeginMaskEdit.Text);
      EndDate := StrToDate(arEndMaskEdit.Text);
      avg_rateReport(Sender);
    end; // конец savg_ratePage

  if toolsPageControl.ActivePage.Name = sChangeOperationsPage then
    begin
      BeginDate := StrToDate(chopBeginMaskEdit.Text);
      EndDate := StrToDate(chopEndMaskEdit.Text);
      ReportHeader := 'Журнал изменений в базе данных ДИСа за отчетный период с '
                      + chopBeginMaskEdit.Text
                      + ' по '
                      + chopEndMaskEdit.Text
                      + ' начиная с '
                      + JournalDateMaskEdit.Text;
         // формируем отчет
      ExportChangeOperations(Sender);
    end; // конец sChangeOperationsPage


  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TtoolsForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
