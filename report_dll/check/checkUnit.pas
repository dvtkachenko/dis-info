unit checkUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, shared_type;

const
  icontract_relPage = 0;
  ino_contract_relPage = 1;
  scontract_relPage = 'contract_relTabSheet';
  sno_contract_relPage = 'no_contract_relTabSheet';
  sContract_relTemplate = 'contract_rel.xlt';

type
  TcheckDataForm = class(TForm)
    checkDataPageControl: TPageControl;
    checkContract_relQuery: TQuery;
    crBeginMaskEdit: TMaskEdit;
    crEndMaskEdit: TMaskEdit;
    contract_relTabSheet: TTabSheet;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    ruleGroupBox: TGroupBox;
    no_contract_relTabSheet: TTabSheet;
    check_no_Contract_relQuery: TQuery;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure checkContract_relReport(Sender: TObject);
    procedure check_no_Contract_relReport(Sender: TObject);
    procedure no_contract_relTabSheetShow(Sender: TObject);
    procedure no_contract_relTabSheetHide(Sender: TObject);
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

procedure TcheckDataForm.FormShow(Sender: TObject);
begin
  crBeginMaskEdit.Text := startDate;
  crEndMaskEdit.Text := DateToStr(Date);
end;

//---------------------------------------------------------------------
// вытаскивает непривязанные операции в БД ДИС98
//---------------------------------------------------------------------
procedure TcheckDataForm.checkContract_relReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..8] of Variant;
     PathToTemplate : string;
     i : integer;
//     ReportHeader : string;
     row : integer;

     { контрольные переменные }
     countUn_rel: integer ;
     //
     debitor_name : string;
     creditor_name : string;
     operation_type : string;
     amount : real;
     pay_date : TDate;
     contract_no : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sContract_relTemplate;
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
     countUn_rel := 0;
     row := 5;

     { просим в базе необходимые данные }
     checkContract_relQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not checkContract_relQuery.Eof do begin
       countUn_rel := countUn_rel + 1;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       debitor_name := checkContract_relQuery.fieldbyname('debitor').asstring;
       creditor_name := checkContract_relQuery.fieldbyname('creditor').asstring;
       operation_type := checkContract_relQuery.fieldbyname('type_name').asstring;
       amount := checkContract_relQuery.fieldbyname('amount').asfloat;
       pay_date := checkContract_relQuery.fieldbyname('pay_date').asdatetime;
       contract_no := checkContract_relQuery.fieldbyname('contract_no').asstring;

       info_row[1] := countUn_rel;
       info_row[2] := debitor_name;
       info_row[3] := creditor_name;
       info_row[4] := operation_type;
       info_row[5] := amount;
       info_row[6] := pay_date;
       info_row[7] := contract_no;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'G' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 8 do info_row[i] := unAssigned;

       row := row + 1;
       checkContract_relQuery.Next;
     end;

   finally
     Excel.free;
     checkContract_relQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

//---------------------------------------------------------------------
// вытаскивает операции, привязанные к договору "без договора" БД ДИС98
//---------------------------------------------------------------------
procedure TcheckDataForm.check_no_Contract_relReport(Sender: TObject);
  Var
     temp: lcid;
     Excel : TExcel;
     cell : string;
     cellFrom : string;
     cellTo : string;
     info_row : array[1..8] of Variant;
     PathToTemplate : string;
     i : integer;
//     ReportHeader : string;
     row : integer;

     { контрольные переменные }
     countUn_rel: integer ;
     //
     debitor_name : string;
     creditor_name : string;
     operation_type : string;
     amount : real;
     pay_date : TDate;
     contract_no : string;

  const
     English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
     Column = 0;
  begin
    temp := GetThreadLocale;
    SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
    PathToTemplate := PathToProgram + '\Template\' + sContract_relTemplate;
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
     countUn_rel := 0;
     row := 5;

     { просим в базе необходимые данные }
     with check_no_Contract_relQuery do begin
       Close;
       Prepare;
       ParamByName('begin_date').asdate := BeginDate;
       ParamByName('end_date').asdate := EndDate;
     end;
     check_no_Contract_relQuery.Open;

  // ---- ---- ----- начало цикла по счетам ----- ----- ----- //
     while not check_no_Contract_relQuery.Eof do begin
       countUn_rel := countUn_rel + 1;
    // ----- ------
       Update;
    // ----- ----- формирование отчета в Excel ------ ------ ------ ------ //
       debitor_name := check_no_Contract_relQuery.fieldbyname('debitor').asstring;
       creditor_name := check_no_Contract_relQuery.fieldbyname('creditor').asstring;
       operation_type := check_no_Contract_relQuery.fieldbyname('type_name').asstring;
       amount := check_no_Contract_relQuery.fieldbyname('amount').asfloat;
       pay_date := check_no_Contract_relQuery.fieldbyname('pay_date').asdatetime;
       contract_no := check_no_Contract_relQuery.fieldbyname('contract_no').asstring;

       info_row[1] := countUn_rel;
       info_row[2] := debitor_name;
       info_row[3] := creditor_name;
       info_row[4] := operation_type;
       info_row[5] := amount;
       info_row[6] := pay_date;
       info_row[7] := contract_no;

       cellFrom := 'A' + IntToStr(row);
       cellTo := 'G' + IntToStr(row);

       Excel.Range[cellFrom,cellTo] := VarArrayOf(info_row);

       for i := 1 to 8 do info_row[i] := unAssigned;

       row := row + 1;
       check_no_Contract_relQuery.Next;
     end;

   finally
     Excel.free;
     check_no_Contract_relQuery.Close;
     SetThreadLocale(Temp);
    end;
end;

procedure TcheckDataForm.sbReportToExcelClick(Sender: TObject);
//Var
//  id : integer;
//  name : string;
//  s : array[0..maxPChar] of Char;
//  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
//  pname := @s;
  BeginDate := StrToDate(crBeginMaskEdit.Text);
  EndDate := StrToDate(crEndMaskEdit.Text);

  if checkDataPageControl.ActivePage.Name = scontract_relPage then
    begin
      ReportHeader := 'Непривязанные операции в базе данных ДИС98';
      checkContract_relReport(Sender);
    end; // конец scontract_relPage

  if checkDataPageControl.ActivePage.Name = sno_contract_relPage then
    begin
      ReportHeader := 'Операции в базе данных ДИС98 привязанные к "без договора" ' +
                      'за период с ' + crBeginMaskEdit.Text +
                      ' по ' + crEndMaskEdit.Text;
      check_no_Contract_relReport(Sender);
    end; // конец ino_contract_relPage

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TcheckDataForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TcheckDataForm.no_contract_relTabSheetShow(Sender: TObject);
begin
  crBeginMaskEdit.Enabled := true;
  crEndMaskEdit.Enabled := true;
end;

procedure TcheckDataForm.no_contract_relTabSheetHide(Sender: TObject);
begin
  crBeginMaskEdit.Enabled := false;
  crEndMaskEdit.Enabled := false;
end;

end.
