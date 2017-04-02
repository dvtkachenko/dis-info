unit zachetUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin, shared_type, word_type;

const
  szachetTabSheet = 'zachetTabSheet';
  imaxSidesCapacity = 30;
  imaxEnt = 10; 
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
  isrc_rowOfData = 2;       // начало данных в исходном файле
  sEndOfData = '';     // если Excel-ячейка равна этому значению
                       // то значит  наступил конец данных в файле

type

  tcontract_info = record
    wdZachetFileName : string;
    full_enterprise_name : string;
    contract_no : string;
    signing_date : TDate;
    sum : real;
    sum_as_text : string;
    dolg_state : string; // кредиторка или дебиторка ?
  end;

  // содержит блок информации о зачете
  TZachet = class
  public
    full_enterprise_name : string;
    zachet_date : TDate;
    wdZachetFileName : string;
    // счетчики кол-ва требований
    count_credit : integer;
    count_debit : integer;
    // общие суммы кредиторских и дебиторских требований
    credit_sum : real;
    debit_sum : real;
    // информация о требованиях кредитора
    credit_side : array[1..imaxSidesCapacity] of tcontract_info;
    // информация о требованиях дебитора
    debit_side : array[1..imaxSidesCapacity] of tcontract_info;
    procedure exportToWord(Var Word : TWord);
  end;

  //
  TExcelToWord = class
  private
    old_lang: lcid;
  public
    constructor Create;
    destructor Destroy; override;
    Excel : TExcel;
    Word : TWord;
    // счетчик кол-ва предприятий в зачете
    count_ent : integer;
    // информация о всех зачетах
    all_zachet : array[1..imaxEnt] of TZachet;
    procedure readFromExcel(src_filename : string);
    procedure all_exportToWord;
  end;

  TzachetForm = class(TForm)
    zachetPageControl: TPageControl;
    zachetTabSheet: TTabSheet;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    fileNameLabel: TLabel;
    source_for_zachetOpenDialog: TOpenDialog;
    openButton: TButton;
    wdopenButton: TButton;
    wdcloseButton: TButton;
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure zachetCreate(Sender: TObject);
    procedure openButtonClick(Sender: TObject);
    procedure wdopenButtonClick(Sender: TObject);
    procedure wdcloseButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    parentConfig : p_config;
    ReportHeader : string;
    PathToProgram : string;
    ExcelToWord : TExcelToWord;
    Word : TWord;
  end;

implementation

uses excel_type;

{$R *.DFM}

{сервисные процедуры}

{-------------------}



//---------------------------------------------------------------------
// выполняет формирование зачета в Word
procedure TZachet.exportToWord;
begin
  //
end;
//---------------------------------------------------------------------
procedure TExcelToWord.Create;
begin
  inherited Create;
  Excel := nil;
  Word := nil;

  old_lang := GetThreadLocale;
  SetThreadLocale(English_Locale);

  try
    Excel := TExcel.Create;
    Word := TWord.Create;
  except
    if Excel <> nil then Excel.Free;
    if Word <> nil then Word.Free;
    SetThreadLocale(old_lang);
  end;
end;
//---------------------------------------------------------------------
procedure TExcelToWord.Destroy;
begin
  if Excel <> nil then Excel.Free;
  if Word <> nil then Word.Free;
  SetThreadLocale(old_lang);
  inherited Destroy;
end;
//---------------------------------------------------------------------
procedure TExcelToWord.readFromExcel(src_filename : string);
Var
  contract_info : tcontract_info;
  row : integer;
  end_of_file : boolean;
  cell : string;
begin
  try
    Excel.AddWorkBook(src_filename);
  except
    raise Exception.Create('Невозможно загрузить Excel');
  end;
  try
    row := isrc_rowOfData;
    end_of_file := false;

    while not end_of_file do begin
      cell := 'A' + IntToStr(row);
      contract_info.wdZachetFileName := Excel.Cell[cell];    // выбираем имя файла зачета  

      // если данные не обнаружены в назначении, в поставщике, в грузополучателе
      // то делаем вывод что достигнут конец файла и мы выходим из цикла
      if ((CoalMessage.dest_coal = sEndOfData) and
          (CoalMessage.coal_sender = sEndOfData) and
          (CoalMessage.cargo_receiver = sEndOfData)) then begin
        end_of_file := true;
        continue;
      end;

      // если данные не обнаружены хотя бы в одном из
      // обязательных полей, то генерируем исключение и выходим
      // из программы
      if ((CoalMessage.dest_coal = sEndOfData) or
          (CoalMessage.coal_sender = sEndOfData) or
          (CoalMessage.amount_free_nds = 0) or
          (CoalMessage.cargo_receiver = sEndOfData)) then begin
        raise Exception.Create('Отсутствуют данные в одном из обязательных полей');
      end;

      // посылаем сообщение CoalMessage объекту Coal
//      AddToSelf(CoalMessage);
      //
      row := row + 1;
    end;
  except
    raise;
  end;

end;
//---------------------------------------------------------------------
procedure TExcelToWord.all_exportToWord;
begin
  //
end;

//---------------------------------------------------------------------
// выполняет формирование зачетов встречных требований
// в Word на основе данных в Excel
//---------------------------------------------------------------------
procedure TzachetForm.zachetCreate(Sender: TObject);
begin
  ExcelToWord := TExcelToWord.Create;







  temp := GetThreadLocale;
  SetThreadLocale(English_Locale);

    Excel := TExcel.Create;
//    PathToTemplate := PathToProgram + '\Template\' + savg_rateTemplate;
    try
      Excel.AddWorkBook(PathToTemplate);
      Excel.Visible := true;
    except
      raise Exception.Create('Невозможно загрузить Excel');
    end;

   try

   finally
     Excel.free;
     SetThreadLocale(Temp);
   end;
end;

procedure TzachetForm.sbReportToExcelClick(Sender: TObject);
//Var
//  id : integer;
//  name : string;
//  s : array[0..maxPChar] of Char;
//  pname : PChar;
begin
//  if mdRadioGroup.ItemIndex = 1 then
  { конструирование запросов }
//  pname := @s;

  if zachetPageControl.ActivePage.Name = szachetTabSheet then
    begin
//      avg_rateReport(Sender);
    end; // конец szachetTabSheet

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TzachetForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TzachetForm.openButtonClick(Sender: TObject);
begin
  if source_for_zachetOpenDialog.Execute then begin
    sbReportToExcel.Enabled := true;
    fileNameLabel.Caption := source_for_zachetOpenDialog.FileName;
  end;
end;

procedure TzachetForm.wdopenButtonClick(Sender: TObject);
Var
  text : WideString;
begin
  try
    Word := Tword.Create;
    Word.OpenDocument('e:\ш.Калиновская-Восточная.doc');
    Word.Visible := true;
    Word.EndKey;
    text := 'Настоящим сообщаем, что в соответствии со ст.217 ' +
            'Гражданского кодекса Украины, с 31.08.2002 г. ' +
            'ДП Корпорации "Индустриальный Союз Донбасса" "Донецкий ' +
            'Индустриальный Союз" прекращены взаимные обязательства ' +
            'зачётом встречных однородных требований, а именно:';
    Word.TypeText(text);
    Word.TypeParagraph;
    text := '1.	Требование ГОАО Шахта "Калиновская-Восточная" ДП ' +
            'ГХК  "Макеевуголь" к  ДП Корпорации "ИСД" "ДИС" на ' +
            'сумму  574 069,84 (пятьсот семьдесят четыре тысячи ' +
            'шестьдесят девять гривень 84 копейки), возникшее в ' +
            'соответствии с договором № ДИС/2-148 псту от 25.03.2002 г.';
    Word.TypeText(text);
  except
    Word.Free;
  end;
end;

procedure TzachetForm.wdcloseButtonClick(Sender: TObject);
begin
  Word.Free;
end;

end.
