unit statisticSaldoCloseUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Menus, DBTables, shared_type, ActnList,
  ComCtrls, Db, Buttons, ToolWin, ImgList, Mask, ComObj, excel_type;

const
  sSaldoCloseTemplate = 'close_saldo.xlt';
  sSaldoClosePageName = 'saldoCloseTabSheet';

type

  TstatisticSaldoCloseForm = class(TForm)
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ExitSpeedButton: TSpeedButton;
    ToolButton1: TToolButton;
    StatisticPageControl: TPageControl;
    saldoCloseTabSheet: TTabSheet;
    Label2: TLabel;
    allCoalSenderQuery: TQuery;
    dis_ibdbDatabase: TDatabase;
    blackListEntQuery: TQuery;
    saldoCloseMaskEdit: TMaskEdit;
    allOtherEntContractQuery: TQuery;
    allCoalEntContractQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure SaldoCloseReport(Sender: TObject);
    procedure SaldoCloseShortReport(Excel : TExcel);
    procedure SaldoCloseDetailReport(Excel : TExcel);
  private
    { Private declarations }
  public
    { Public declarations }
    PathToProgram : string;
    ReportHeader : string;
    saldoCloseDate : TDateTime;
  end;


var
  statisticSaldoCloseForm: TstatisticSaldoCloseForm;

implementation

uses serviceDataUnit;

{$R *.DFM}

function GetEnterprise(Var id:integer; pname: PChar) : integer; external 'service.dll' name 'GetEnterprise';
function GetContract(id:integer;Var pcontract_no: PChar) : integer; external 'service.dll' name 'GetContract';

{сервисные процедуры}

{-------------------}

procedure TstatisticSaldoCloseForm.FormShow(Sender: TObject);
begin
  saldoCloseMaskEdit.Text := DateToStr(Date);
end;

//---------------------------------------------------------------------
// формирование статистики по всем поставщикам угля
// по которым возможно закрытие взаимных задолженностей
// путем прогонки денег , векселей или соглашения о зачете
// встречных требований
//---------------------------------------------------------------------
procedure TstatisticSaldoCloseForm.SaldoCloseReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     PathToTemplate : string;
     i : integer;
     blackList : string;
     id : integer;

     { контрольные переменные }
     countEnterpr : integer ;

  const
    English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sSaldoCloseTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try
     saldoCloseDate := StrToDate(saldoCloseMaskEdit.Text);
     blackList := '';

     { формируем список договоров по которым необходимо
       закрыться  }
     blackListEntQuery.Open;
     while not blackListEntQuery.Eof do begin
       id := blackListEntQuery.fieldbyname('enterpr_id').asinteger;
       blackList := blackList + IntToStr(id) + ',';
       blackListEntQuery.Next;
     end;
     // чтобы не выдирать запятую из строки
     // поставим после id предприятия = 0
     blackList := blackList + '0';
     //
     with allCoalSenderQuery do begin
       Close;
       SQL.Clear;
       SQL.Add('select distinct enterprise_id, enterprise_name');
       SQL.Add('from all_coal_saldo_date(:saldo_date)');
       SQL.Add('where enterprise_id not in(');
       SQL.Add(blackList);
       SQL.Add(')');
       SQL.Add('order by enterprise_name');
       ParamByName('saldo_date').asdate := saldoCloseDate;
       Open;
     end;

     // формирование свернутого отчета по сальдо
     Excel.SelectWorkSheet('saldo');
     SaldoCloseShortReport(Excel);

     // формирование развернутого отчета по сальдо
//     Excel.SelectWorkSheet('detail');
//     SaldoCloseDetailReport(Excel);

   finally
     Excel.free;
     blackListEntQuery.Close;
     allCoalSenderQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

//---------------------------------------------------------------------
// формирование свернутой статистики (лист saldo)
// по всем поставщикам угля
// по которым возможно закрытие взаимных задолженностей
// путем прогонки денег , векселей или соглашения о зачете
// встречных требований
//---------------------------------------------------------------------
procedure TstatisticSaldoCloseForm.SaldoCloseShortReport(Excel : TExcel);
  Var
     temp: lcid;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..12] of Variant;
     row, rowCoal, rowOther : integer;
     i : integer;
     ent_id : integer;
     debit, credit : real;
     flag : boolean;

     // буферные переменные
     countEnterpr : integer;
     enterpr_name : string;
     coal_contract : string;
     coal_contract_date : TDate;
     coal_role_name : string;
     coal_contract_saldo : real;
     other_contract : string;
     other_contract_date : TDate;
     other_role_name : string;
     other_contract_saldo : real;
     full_saldo : real;

  begin
   try
     { контрольные переменные }
     row := 2;
     cell := 'A' + IntToStr(row);
     ReportHeader := 'Сальдо по всем договорам поставки угля на ';
     ReportHeader := ReportHeader + DateToStr(saldoCloseDate);
     Excel.Cell[cell] := ReportHeader;

     countEnterpr := 0;
     row := 7;
     flag := false;

  // ---- ---- ----- начало цикла по договорам ----- ----- ----- //
     while not allCoalSenderQuery.Eof do begin
       ent_id := allCoalSenderQuery.fieldbyname('enterprise_id').asinteger;
       enterpr_name := allCoalSenderQuery.fieldbyname('enterprise_name').asstring;

       // вытаскиваем сальдо по угольным договорам
       with allCoalEntContractQuery do begin
         Close;
         ParamByName('ent_id').asinteger := ent_id;
         ParamByName('saldo_date').asdate := saldoCloseDate;
         Open;
       end;

       // вытаскиваем сальдо по прочим договорам
       with allOtherEntContractQuery do begin
         Close;
         ParamByName('ent_id').asinteger := ent_id;
         ParamByName('saldo_date').asdate := saldoCloseDate;
         Open;
       end;

       rowCoal := row;  // запоминаем значение строки для угольных контрактов
       rowOther := row; // запоминаем значение строки для прочих контрактов

    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       //  прочие договора
       while not allOtherEntContractQuery.Eof do begin
         other_contract := allOtherEntContractQuery.fieldbyname('contract').asstring;
         other_contract_date := allOtherEntContractQuery.fieldbyname('signing_date').asdatetime;
         other_role_name := allOtherEntContractQuery.fieldbyname('role_name').asstring;
         debit := allOtherEntContractQuery.fieldbyname('debit').asfloat;
         credit := allOtherEntContractQuery.fieldbyname('credit').asfloat;
         other_contract_saldo := debit - credit;;
         if (debit <> 0) or (credit <> 0) then begin

           info_row[1] := other_contract;
           info_row[2] := other_contract_date;
           info_row[3] := other_role_name;
           info_row[4] := other_contract_saldo;

           cellFrom := 'G' + IntToStr(rowOther);
           cellTo := 'J' + IntToStr(rowOther);
           Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
           // устанавливаем признак наличия какого-либо прочего договора
           flag := true;

           rowOther := rowOther + 1;
         end;
         allOtherEntContractQuery.Next;
       end;
       //
       if flag then begin
         countEnterpr := countEnterpr + 1;
         info_row[1] := countEnterpr;
         info_row[2] := enterpr_name;

         cellFrom := 'A' + IntToStr(row);
         cellTo := 'B' + IntToStr(row);
         Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);
       end;

       // договора оплаты углей
       while not allCoalEntContractQuery.Eof do begin
         coal_contract := allCoalEntContractQuery.fieldbyname('contract').asstring;
         coal_contract_date := allCoalEntContractQuery.fieldbyname('signing_date').asdatetime;
         coal_role_name := allCoalEntContractQuery.fieldbyname('role_name').asstring;
         debit := allCoalEntContractQuery.fieldbyname('debit').asfloat;
         credit := allCoalEntContractQuery.fieldbyname('credit').asfloat;
         coal_contract_saldo := debit - credit;;

         if flag then begin
           info_row[1] := coal_contract;
           info_row[2] := coal_contract_date;
           info_row[3] := coal_role_name;
           info_row[4] := coal_contract_saldo;

           cellFrom := 'C' + IntToStr(rowCoal);
           cellTo := 'F' + IntToStr(rowCoal);
           Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

           rowCoal := rowCoal + 1;
         end;
         allCoalEntContractQuery.Next;
       end;

       for i := 1 to 12 do info_row[i] := unAssigned;

       if flag then begin
         if rowCoal < rowOther then
           row := rowOther
         else
           row := rowCoal;

         row := row + 1;
       end;

       // сбрасываем флаг признака наличия прочего договора
       flag := false;

       allCoalSenderQuery.Next;
     end; // конец цикла по предприятиям allCoalSenderQuery

   finally
     allCoalEntContractQuery.Close;
     allOtherEntContractQuery.Close;
   end;
end;

//---------------------------------------------------------------------
// формирование детальной статистики (лист detail)
// по всем поставщикам угля
// по которым возможно закрытие взаимных задолженностей
// путем прогонки денег , векселей или соглашения о зачете
// встречных требований
//---------------------------------------------------------------------
procedure TstatisticSaldoCloseForm.SaldoCloseDetailReport(Excel : TExcel);
begin
  //
end;

procedure TstatisticSaldoCloseForm.sbReportToExcelClick(Sender: TObject);
begin
  //
  if StatisticPageControl.ActivePage.Name = sSaldoClosePageName then
    SaldoCloseReport(Sender);
end;

/////////////////////////////////////////////////////////////////////////
procedure TstatisticSaldoCloseForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

end.
