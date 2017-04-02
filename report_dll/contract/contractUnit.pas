unit contractUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin;

const
  scontractPage = 'contractTabSheet';
  scontract_infoTemplate = 'contract_info.xlt';
  sall_dolgiContractPage = 'all_dolgiContractTabSheet';
  sall_dolgiContractTemplate = 'all_dolgiContract.xlt';
type
  TContractForm = class(TForm)
    ContractPageControl: TPageControl;
    allContractOldQuery: TQuery;
    contractTabSheet: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    allContractQuery: TQuery;
    contractBeginMaskEdit: TMaskEdit;
    contractEndMaskEdit: TMaskEdit;
    changeTabSheet: TTabSheet;
    Label3: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    journalDateMaskEdit: TMaskEdit;
    Label9: TLabel;
    changeBeginMaskEdit: TMaskEdit;
    changeEndMaskEdit: TMaskEdit;
    all_dolgiContractTabSheet: TTabSheet;
    all_dolgiContractEndMaskEdit: TMaskEdit;
    Label5: TLabel;
    all_dolgiContractQuery: TQuery;
    dolgi_prihodQuery: TQuery;
    dolgi_rashodQuery: TQuery;
    count_dolgi_prihodQuery: TQuery;
    count_dolgi_rashodQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure exportContractInfo(Sender: TObject);
    procedure export_all_dolgiContract(Sender: TObject);
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


{сервисные процедуры}

{-------------------}

procedure TContractForm.FormShow(Sender: TObject);
begin
  contractBeginMaskEdit.Text := startDate;
  contractEndMaskEdit.Text := DateToStr(Date);
{ changeBeginMaskEdit.Text := startDate;
  changeEndMaskEdit.Text := DateToStr(Date);
  journalDateMaskEdit.Text := DateToStr(Date-1);}
end;

//---------------------------------------------------------------------
// формирует отчет по договорам за указанный период
//---------------------------------------------------------------------
procedure TContractForm.exportContractInfo(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..14] of Variant;
     PathToTemplate : string;
     i : integer;
     row : integer;

     { контрольные переменные }
     countContract : integer ;
     //
     contract_no : string;
     signing_date : TDate;
     contract_sum : real;
     to_base_date : TDate;
     contract_char_id : string;
     contract_type : string;
     count_subdivision : integer;
     first_subdivision_name : string;
     enterprise_name : string;
     saldo_begin : real;
     operations_type : string;
     debit : real;
     credit : real;
     saldo_end : real;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    try
      Excel := TExcel.Create;
    except
      raise Exception.Create('Невозможно создать OLE - объект');
    end;

    PathToTemplate := PathToProgram + '\Template\' + scontract_infoTemplate;
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

     BeginDate := StrToDate(contractBeginMaskEdit.Text);
     EndDate := StrToDate(contractEndMaskEdit.Text);

     { инициализируем  контрольные переменные }
     countContract := 0;
     row := 6;

     { просим в базе информацию по всем договорам }
     with allContractQuery do begin
       Close;
       Prepare;
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
     end;
     allContractQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not allContractQuery.Eof do begin
       countContract := countContract + 1;

       contract_no := allContractQuery.fieldbyname('contract_no').asstring;
       signing_date := allContractQuery.fieldbyname('signing_date').asdatetime;
       contract_sum := allContractQuery.fieldbyname('contract_sum').asfloat;
//       to_base_date := allContractQuery.fieldbyname('to_base_date').asdatetime;
       contract_char_id := allContractQuery.fieldbyname('contract_char_id').asstring;
       contract_type := allContractQuery.fieldbyname('contract_type').asstring;
//       count_subdivision := allContractQuery.fieldbyname('count_subdivision').asinteger;
//       first_subdivision_name := allContractQuery.fieldbyname('first_subdivision_name').asstring;
       enterprise_name := allContractQuery.fieldbyname('enterprise_name').asstring;
       saldo_begin := allContractQuery.fieldbyname('saldo_begin').asfloat;
       operations_type := allContractQuery.fieldbyname('operations_type').asstring;
       debit := allContractQuery.fieldbyname('debit').asfloat;
       credit := allContractQuery.fieldbyname('credit').asfloat;
       saldo_end := allContractQuery.fieldbyname('saldo_end').asfloat;

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       info_row[1] := contract_no;
       info_row[2] := signing_date;
       info_row[3] := contract_sum;
//       info_row[4] := to_base_date;
       info_row[5] := contract_char_id;
       info_row[6] := contract_type;
//       info_row[7] := count_subdivision;
//       info_row[8] := first_subdivision_name;
       info_row[9] := enterprise_name;
       info_row[10] := saldo_begin;
       info_row[11] := operations_type;
       info_row[12] := debit;
       info_row[13] := credit;
       info_row[14] := saldo_end;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'N' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       for i := 1 to 14 do info_row[i] := unAssigned;

       row := row + 1;

       allContractQuery.Next;
     end;

   finally
     Excel.free;
     allContractQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

//---------------------------------------------------------------------
// формирует отчет по всем договорам за весь период деятельности ДИСа
// данный отчет применяется для списания задолженности
//---------------------------------------------------------------------
procedure TContractForm.export_all_dolgiContract(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..22] of Variant;
     PathToTemplate : string;
     i : integer;
     row : integer;
     exit_flag : boolean;  // флаг выхода из цикла
     prihod_flag : boolean;  // флаг разрешения чтения строки ПРИХОДА
     rashod_flag : boolean;  // флаг разрешения чтения строки РАСХОДА

     // вспомогательные переменные
     count_rec_prihod : integer;
     count_rec_rashod : integer;

     prihod_type_oper : string;
     prihod_date_oper : TDate;
     prihod_sum_oper : real;
     rashod_type_oper : string;
     rashod_date_oper : TDate;
     rashod_sum_oper : real;

     dohod_sum : real;
     rashod_sum : real;


     { контрольные переменные }
     countContract : integer ;
     //
     enterprise_id : integer;
     enterprise_name : string;
     contract_id : integer;
     contract_no : string;
     contract_type : string;
     signing_date : TDate;
     enterprise_role : string;
     contract_sum : real;
     debit : real;
     credit : real;
     saldo : real;
     //
     debit_type_oper : string;
     debit_date_oper : TDate;
     debit_sum_oper : real;
     credit_type_oper : string;
     credit_date_oper : TDate;
     credit_sum_oper : real;
     //
     debit_delta_time : TDate;
     credit_delta_time : TDate;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    try
      Excel := TExcel.Create;
    except
      raise Exception.Create('Невозможно создать OLE - объект');
    end;

    PathToTemplate := PathToProgram + '\Template\' + sall_dolgiContractTemplate;
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

     EndDate := StrToDate(all_dolgiContractEndMaskEdit.Text);

     { инициализируем  контрольные переменные }
     countContract := 0;
     row := 6;

     { просим в базе информацию по всем договорам }
     with all_dolgiContractQuery do begin
       Close;
       Prepare;
       ParamByName('pay_date').asdate := EndDate;
     end;
     all_dolgiContractQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not all_dolgiContractQuery.Eof do begin
       countContract := countContract + 1;
       //
       enterprise_id := all_dolgiContractQuery.fieldbyname('enterprise_id').asinteger;
       enterprise_name := all_dolgiContractQuery.fieldbyname('enterprise_name').asstring;
       contract_id := all_dolgiContractQuery.fieldbyname('contract_id').asinteger;
       contract_no := all_dolgiContractQuery.fieldbyname('contract_no').asstring;
       contract_type := all_dolgiContractQuery.fieldbyname('contract_type').asstring;
       signing_date := all_dolgiContractQuery.fieldbyname('signing_date').asdatetime;
       enterprise_role := all_dolgiContractQuery.fieldbyname('role_name').asstring;
       contract_sum := all_dolgiContractQuery.fieldbyname('contract_sum').asfloat;
       debit := all_dolgiContractQuery.fieldbyname('debit').asfloat;
       credit := all_dolgiContractQuery.fieldbyname('credit').asfloat;
       saldo := debit-credit;
       //
       { просим в базе информацию по договору ПРИХОД }
       with dolgi_prihodQuery do begin
         Close;
         Prepare;
         ParamByName('ent_id').asinteger := enterprise_id;
         ParamByName('contract_no').asstring := contract_no;
         ParamByName('pay_date').asdate := EndDate;
       end;
       dolgi_prihodQuery.Open;
       { считаем кол-во записей }
{       with count_dolgi_prihodQuery do begin
         Close;
         Prepare;
         ParamByName('ent_id').asinteger := enterprise_id;
         ParamByName('contract_no').asstring := contract_no;
         ParamByName('pay_date').asdate := EndDate;
       end;
       count_dolgi_prihodQuery.Open;
       count_rec_prihod := count_dolgi_prihodQuery.fieldbyname('count_rec').asinteger;;
}
       { просим в базе информацию по договору РАСХОД }
       with dolgi_rashodQuery do begin
         Close;
         Prepare;
         ParamByName('ent_id').asinteger := enterprise_id;
         ParamByName('contract_no').asstring := contract_no;
         ParamByName('pay_date').asdate := EndDate;
       end;
       dolgi_rashodQuery.Open;
       { считаем кол-во записей }
{       with count_dolgi_rashodQuery do begin
         Close;
         Prepare;
         ParamByName('ent_id').asinteger := enterprise_id;
         ParamByName('contract_no').asstring := contract_no;
         ParamByName('pay_date').asdate := EndDate;
       end;
       count_dolgi_rashodQuery.Open;
       count_rec_rashod := count_dolgi_rashodQuery.fieldbyname('count_rec').asinteger;;
}
       // начало цикла расчета даты списания
       exit_flag := false;
       dohod_sum := 0;
       rashod_sum := 0;
       prihod_flag := true;
       rashod_flag := true;
       while not exit_flag do begin

         // выбираем операцию ПРИХОДА
         if prihod_flag then begin
           prihod_type_oper := dolgi_prihodQuery.fieldbyname('type_name').asstring;
           prihod_date_oper := dolgi_prihodQuery.fieldbyname('pay_date').asdatetime;
           prihod_sum_oper := dolgi_prihodQuery.fieldbyname('amount').asfloat;
           dolgi_prihodQuery.Next;
           dohod_sum := dohod_sum + prihod_sum_oper;
         end;

         // выбираем операцию РАСХОДА
         if rashod_flag then begin
           rashod_type_oper := dolgi_rashodQuery.fieldbyname('type_name').asstring;
           rashod_date_oper := dolgi_rashodQuery.fieldbyname('pay_date').asdatetime;
           rashod_sum_oper := dolgi_rashodQuery.fieldbyname('amount').asfloat;
           dolgi_rashodQuery.Next;
           rashod_sum := rashod_sum + rashod_sum_oper;
         end;

         if dolgi_prihodQuery.Eof then prihod_flag := false;
         if dolgi_rashodQuery.Eof then rashod_flag := false;

         if ((dohod_sum = rashod_sum) and
             ((dolgi_prihodQuery.Eof) or
              (dolgi_rashodQuery.Eof))) then begin
           debit_type_oper := prihod_type_oper;
           debit_date_oper := prihod_date_oper;
           debit_sum_oper := prihod_sum_oper;
           credit_type_oper := rashod_type_oper;
           credit_date_oper := rashod_date_oper;
           credit_sum_oper := rashod_sum_oper;
           exit_flag := true;
           continue;
         end;

         if ((dohod_sum < rashod_sum) and
             (dolgi_prihodQuery.Eof)) then begin
           debit_type_oper := prihod_type_oper;
           debit_date_oper := prihod_date_oper;
           debit_sum_oper := prihod_sum_oper;
           credit_type_oper := rashod_type_oper;
           credit_date_oper := rashod_date_oper;
           credit_sum_oper := rashod_sum_oper;
           exit_flag := true;
           continue;
         end;

         if ((dohod_sum > rashod_sum) and
             (dolgi_rashodQuery.Eof)) then begin
           debit_type_oper := prihod_type_oper;
           debit_date_oper := prihod_date_oper;
           debit_sum_oper := prihod_sum_oper;
           credit_type_oper := rashod_type_oper;
           credit_date_oper := rashod_date_oper;
           credit_sum_oper := rashod_sum_oper;
           exit_flag := true;
           continue;
         end;

         if ((dohod_sum = rashod_sum) and
             (not (dolgi_prihodQuery.Eof) and
              not (dolgi_rashodQuery.Eof))) then begin
           prihod_flag := true;
           rashod_flag := true;
           continue;
         end;

         if ((dohod_sum < rashod_sum) and
             not (dolgi_prihodQuery.Eof)) then begin
           prihod_flag := true;
           rashod_flag := false;
           continue;
         end;

         if ((dohod_sum > rashod_sum) and
             not (dolgi_rashodQuery.Eof)) then begin
           prihod_flag := false;
           rashod_flag := true;
           continue;
         end;

       end;  // конец цикла расчета даты списания

     // вычисляем дельту в месяцах
     debit_delta_time := (EndDate-debit_date_oper)/30;
     credit_delta_time := (EndDate-credit_date_oper)/30;

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       info_row[1] := countContract;
       info_row[2] := enterprise_name;
       info_row[3] := contract_no;
       info_row[4] := contract_type;
       info_row[5] := signing_date;
       info_row[6] := enterprise_role;
       info_row[7] := contract_sum;
       info_row[9] := debit;
       info_row[10] := credit;
       info_row[11] := saldo;

       if (debit_sum_oper = 0) then begin
         info_row[13] := ' ';
         info_row[14] := ' ';
         info_row[15] := ' ';
       end
       else begin
         info_row[13] := debit_type_oper;
         info_row[14] := debit_date_oper;
         info_row[15] := debit_sum_oper;
       end;

       if (credit_sum_oper = 0) then begin
         info_row[16] := ' ';
         info_row[17] := ' ';
         info_row[18] := ' ';
       end
       else begin
         info_row[16] := credit_type_oper;
         info_row[17] := credit_date_oper;
         info_row[18] := credit_sum_oper;
       end;

       if (debit_sum_oper = 0) then
         info_row[19] := ' '
       else
         info_row[19] := debit_delta_time;

       if (credit_sum_oper = 0) then
         info_row[20] := ' '
       else
         info_row[20] := credit_delta_time;

       info_row[21] := enterprise_id;
       info_row[22] := contract_id;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'V' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       for i := 1 to 22 do info_row[i] := unAssigned;

       row := row + 1;

       all_dolgiContractQuery.Next;
     end;

   finally
     Excel.free;
     all_dolgiContractQuery.Close;
     dolgi_rashodQuery.Close;
     dolgi_prihodQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

//---------------------------------------------------------------
procedure TContractForm.sbReportToExcelClick(Sender: TObject);
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
  BeginDate := StrToDate(contractBeginMaskEdit.Text);
  EndDate := StrToDate(contractEndMaskEdit.Text);

  if ContractPageControl.ActivePage.Name = scontractPage then
       begin
         ReportHeader := 'Информация о договорах за период с ' +
                  contractBeginMaskEdit.Text + ' по ' + contractEndMaskEdit.Text;

         // формируем отчет
         exportContractInfo(Sender);
       end; // конец scontractPage

  if ContractPageControl.ActivePage.Name = sall_dolgiContractPage then
       begin
         ReportHeader := 'Информация для списания просроченной '
                         + 'задолженности по состоянию на ' +
                         all_dolgiContractEndMaskEdit.Text;

         // формируем отчет
         export_all_dolgiContract(Sender);
       end; // конец sall_dolgiContractPage

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TContractForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
