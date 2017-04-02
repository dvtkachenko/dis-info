unit coalUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin;

const
  ispacingCoalTabSheet = 0;
  sPriplSkidki = '';   // то чему должно равняться TCoalMessage.sort_name
                       // для того чтобы сумма разбивалась пропорционально
                       // по тоннам в классе  TCoalReceiver
  sdest_filename = 'coal_for_balans.xlt'; // имя файла отчета
  sSheetNameOfsrcData = 'coal';  // имя листа с исходными данными

  idest_rowOfData = 2;       // начало данных в файле отчета
  sdest_colSender = 'A';
  sdest_colCoalSort = 'C';
  sdest_colCoalQnty = 'D';
  sdest_colSum_free_nds = 'E';
  sdest_colCoalReceiver = 'G';
  sdest_colNotes = 'H';
  sdest_nameItogo = 'Итого:';  // наименование строки итогов
                               // для вывода в отчет

  isrc_rowOfData = 2;       // начало данных в исходном файле
  sEndOfData = '';     // если Excel-ячейка равна этому значению
                       // то значит  наступил конец данных в файле

//  sndsReportTemplate = 'nds_report.xlt';
//  iMaxDept = 10;

type
  // содержит информацию из строки Excel
  TCoalMessage = record
    dest_coal : string;       // назначение углей
    coal_sender : string;     // поставщик углей
    sort_name : string;       // марка угля
    qnty : real;
    amount_free_nds : real;
    amount_whith_nds : real;
    nds : real;
    cargo_receiver : string;  // грузополучатель
    // вспомогательные переменные
    avg_pripl_skidki_free_nds : real; // переменная содержащая рассчитанное среднепропорциональное
                                      // значение приплат/скидок
    readFlag : boolean;       // устанавливается в true если сообщение
                              // получено и обработано
  end;

  // содержит блок информации о марке угля его количестве и сумме
  TCoalSort = class
  private
    sort_name : string;
    qnty : real;
    amount_free_nds : real;
    amount_whith_nds : real;
    nds : real;
    // переменная содержащая среднепропорциональное
    // значение приплат/скидок
    avg_pripl_skidki_free_nds : real;
  public
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
    procedure GetInfo(Var CoalMessage : TCoalMessage);
    procedure Calc(avg:real); // распределение среднепропорционально
    constructor Create(Var CoalMessage : TCoalMessage);
    destructor Destroy; override;
  end;

  //
  TCoalReceiver = class
  private
    CoalSortList : TList;
  public
    dest_coal : string;       // назначение углей
    cargo_receiver : string;  // грузополучатель
    pripl_skidki_free_nds : real;      // переменная содержащая значение
                              // которое необходимо распределить
                              // среднепропорционально кол-ву на все эл-ты
                              // содержащиеся в CoalSortList
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
    procedure GetInfo(Var _CoalSortList : TList);
    procedure Calc; // распределение среднепропорционально
    constructor Create(Var CoalMessage : TCoalMessage);
    destructor Destroy; override;
  end;

  // весь уголь от поставщика
  TCoalSender = class
  private
    CoalReceiverList : TList;
  public
    coal_sender : string;     // поставщик углей
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
    procedure GetInfo(Var _CoalReceiverList : TList);
    procedure Calc; // распределение среднепропорционально
    constructor Create(Var CoalMessage : TCoalMessage);
    destructor Destroy; override;
  end;

  // весь уголь
  TCoal = class
  private
    CoalSenderList : TList;
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
  public
    procedure Read(src_filename : string); // процедура чтения данных из Excel-файла
    procedure Calc; // распределение среднепропорционально
    procedure Report(dest_filename : string); // процедура формирования отчета в Excel
    constructor Create;
    destructor Destroy; override;
  end;
  //
  TCoalReportForm = class(TForm)
    mainPageControl: TPageControl;
    spacingCoalTabSheet: TTabSheet;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbReportToExcel: TSpeedButton;
    ToolButton1: TToolButton;
    ExitSpeedButton: TSpeedButton;
    coalOpenDialog: TOpenDialog;
    fileNameLabel: TLabel;
    OpenButton: TButton;
    procedure CreatCoalReport;
    procedure ExitSpeedButtonClick(Sender: TObject);
    procedure sbReportToExcelClick(Sender: TObject);
    procedure OpenButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    ReportHeader : string;
    PathToProgram : string;   // путь к файлу программы
    Coal : TCoal;   //  содержит уголь из файла Excel в структурированном виде 
  end;

implementation

uses excel_type;

{$R *.DFM}

{сервисные процедуры}

{-------------------}

// ---------------------------------------
// --- реализация методов класса TCoalSort
// ---------------------------------------
constructor TCoalSort.Create(Var CoalMessage : TCoalMessage);
begin
  inherited Create;
  // инициализация объекта
  sort_name := CoalMessage.sort_name;
  qnty := CoalMessage.qnty;
  amount_free_nds := CoalMessage.amount_free_nds;
  amount_whith_nds := CoalMessage.amount_whith_nds;
  nds := CoalMessage.nds;
  avg_pripl_skidki_free_nds := 0;
  CoalMessage.readFlag := true; // установка флага получения сообщения
end;
//---------------------------------------
procedure TCoalSort.AddToSelf(Var CoalMessage : TCoalMessage);
begin
  // если сообщение уже было ранее прочитано то выходим из процедуры
  if (CoalMessage.readFlag) then exit;

  // если марка угля в сообщении соответствует марке
  // в данном классе то происходит добавление данных к данным класса
  // т.е. означает что сообщение получено
  if (sort_name = CoalMessage.sort_name) then begin
    qnty := qnty + CoalMessage.qnty;
    amount_free_nds := amount_free_nds + CoalMessage.amount_free_nds;
    amount_whith_nds := amount_whith_nds + CoalMessage.amount_whith_nds;
    nds := nds + CoalMessage.nds;
    CoalMessage.readFlag := true; // установка флага получения сообщения
  end;
end;
//---------------------------------------
procedure TCoalSort.GetInfo(Var CoalMessage : TCoalMessage);
begin
  CoalMessage.sort_name := sort_name;
  CoalMessage.qnty := qnty;
  CoalMessage.amount_free_nds := amount_free_nds;
  CoalMessage.amount_whith_nds := amount_whith_nds;
  CoalMessage.nds := nds;
  CoalMessage.avg_pripl_skidki_free_nds := avg_pripl_skidki_free_nds;
end;
//---------------------------------------
procedure TCoalSort.Calc(avg:real);
begin
  avg_pripl_skidki_free_nds := avg; 
end;
//---------------------------------------
destructor TCoalSort.Destroy;
begin
  // код программы ....
  inherited Destroy;
end;

// ---------------------------------------
// --- реализация методов класса TCoalReceiver
// ---------------------------------------
constructor TCoalReceiver.Create(Var CoalMessage : TCoalMessage);
begin
  inherited Create;
  dest_coal := CoalMessage.dest_coal;       // назначение углей
  cargo_receiver := CoalMessage.cargo_receiver;  // грузополучатель
  pripl_skidki_free_nds := 0;
  CoalSortList := TList.Create;
  // если в первой строке окажутся приплаты скидки
  if (CoalMessage.sort_name = sPriplSkidki) then begin
    pripl_skidki_free_nds := CoalMessage.amount_free_nds;
    CoalMessage.readFlag := true;
  end
  else begin
    CoalSortList.Add(TCoalSort.Create(CoalMessage));
  end;
end;
//---------------------------------------
procedure TCoalReceiver.AddToSelf(Var CoalMessage : TCoalMessage);
Var
  i : integer;
  CoalSort : TCoalSort;
begin
  // если сообщение уже было ранее прочитано то выходим из процедуры
  if (CoalMessage.readFlag) then exit;
  // если сообщение нашло своего получателя, то ...
  if ((dest_coal = CoalMessage.dest_coal) and
      (cargo_receiver = CoalMessage.cargo_receiver)) then begin
     // если в строке есть приплаты скидки
     if (CoalMessage.sort_name = sPriplSkidki) then begin
       pripl_skidki_free_nds := pripl_skidki_free_nds + CoalMessage.amount_free_nds;
       CoalMessage.readFlag := true;
     end
     else begin
       // если список еще пустой
//       if (CoalSortList.Count = 0) then
//         CoalSortList.Add(TCoalSort.Create(CoalMessage));
       // обход всего списка и поиск получателя сообщения
       for i :=1 to CoalSortList.Count do begin
         CoalSort := CoalSortList.Items[i-1];
         CoalSort.AddToSelf(CoalMessage);
         // если сообщение уже было прочитано то выходим из цикла
         if (CoalMessage.readFlag) then break;
       end;
       // добавление еще одного элемента в список если
       // сообщение не получено
       if not (CoalMessage.readFlag) then begin
         CoalSortList.Add(TCoalSort.Create(CoalMessage));
       end;
     end;
  end;
end;
//---------------------------------------
procedure TCoalReceiver.GetInfo(Var _CoalSortList : TList);
begin
   _CoalSortList := CoalSortList;
end;
//---------------------------------------
procedure TCoalReceiver.Calc;
Var
  i : integer;
  all_qnty : real;
  avg : real;
  CoalMessage : TCoalMessage;
  CoalSort : TCoalSort;
begin
  all_qnty := 0;
  // обход всего списка и подсчет общего количества
  // угольного концентрата в данном списке
  for i :=1 to CoalSortList.Count do begin
    CoalSort := CoalSortList.Items[i-1];
    CoalSort.GetInfo(CoalMessage);
    all_qnty := all_qnty + CoalMessage.qnty;
  end;
  // обход всего списка и расчет среднепропорциональным
  // методом приплат/скидок
  for i :=1 to CoalSortList.Count do begin
    CoalSort := CoalSortList.Items[i-1];
    CoalSort.GetInfo(CoalMessage);
    avg := CoalMessage.qnty/all_qnty * pripl_skidki_free_nds;
    // заносим приплаты/скидки
    CoalSort.Calc(avg);
  end;
end;
//---------------------------------------
destructor TCoalReceiver.Destroy;
Var
  i : integer;
  CoalSort : TCoalSort;
begin
  // обход всего списка и удаление всех элементов списка
  // из памяти
  for i :=1 to CoalSortList.Count do begin
    CoalSort := CoalSortList.Items[i-1];
    CoalSort.Free;
  end;
  inherited Destroy;
end;

// ---------------------------------------
// --- реализация методов класса TCoalSender
// ---------------------------------------
constructor TCoalSender.Create(Var CoalMessage : TCoalMessage);
begin
  inherited Create;
  coal_sender := CoalMessage.coal_sender;       // поставщик углей
  CoalReceiverList := TList.Create;
  CoalReceiverList.Add(TCoalReceiver.Create(CoalMessage));
end;
//---------------------------------------
procedure TCoalSender.AddToSelf(Var CoalMessage : TCoalMessage);
Var
  i : integer;
  CoalReceiver : TCoalReceiver;
begin
  // если сообщение уже было ранее прочитано то выходим из процедуры
  if (CoalMessage.readFlag) then exit;
  // если сообщение нашло своего получателя, то ...
  if (coal_sender = CoalMessage.coal_sender) then begin
    // обход всего списка и поиск получателя сообщения
    for i :=1 to CoalReceiverList.Count do begin
      CoalReceiver := CoalReceiverList.Items[i-1];
      CoalReceiver.AddToSelf(CoalMessage);
      // если сообщение уже было прочитано то выходим из цикла
      if (CoalMessage.readFlag) then break;
    end;
    // добавление еще одного элемента в список если
    // сообщение не получено
    if not (CoalMessage.readFlag) then begin
      CoalReceiverList.Add(TCoalReceiver.Create(CoalMessage));
    end;
  end;
end;
//---------------------------------------
procedure TCoalSender.GetInfo(Var _CoalReceiverList : TList);
begin
  _CoalReceiverList := CoalReceiverList;
end;
//---------------------------------------
procedure TCoalSender.Calc;
Var
  i : integer;
  CoalReceiver : TCoalReceiver;
begin
  // обход всего списка и вызов метода Calc
  for i :=1 to CoalReceiverList.Count do begin
    CoalReceiver := CoalReceiverList.Items[i-1];
    CoalReceiver.Calc;
  end;
end;
//---------------------------------------
destructor TCoalSender.Destroy;
Var
  i : integer;
  CoalReceiver : TCoalReceiver;
begin
  // обход всего списка и удаление всех элементов списка
  // из памяти
  for i :=1 to CoalReceiverList.Count do begin
    CoalReceiver := CoalReceiverList.Items[i-1];
    CoalReceiver.Free;
  end;
  // код программы ....
  inherited Destroy;
end;

// ---------------------------------------
// --- реализация методов класса TCoal
// ---------------------------------------
constructor TCoal.Create;
begin
  inherited Create;
  CoalSenderList := TList.Create;
end;
//---------------------------------------
procedure TCoal.AddToSelf(Var CoalMessage : TCoalMessage);
Var
  i : integer;
  CoalSender : TCoalSender;
begin
  // если сообщение уже было ранее прочитано то выходим из процедуры
  if (CoalMessage.readFlag) then exit;
  // обход всего списка и поиск получателя сообщения
  for i :=1 to CoalSenderList.Count do begin
    CoalSender := CoalSenderList.Items[i-1];
    CoalSender.AddToSelf(CoalMessage);
    // если сообщение уже было прочитано то выходим из цикла
    if (CoalMessage.readFlag) then break;
  end;
  // добавление еще одного элемента в список если
  // сообщение не получено
  if not (CoalMessage.readFlag) then begin
    CoalSenderList.Add(TCoalSender.Create(CoalMessage));
  end;
end;
//---------------------------------------
procedure TCoal.Report(dest_filename : string);
Var
  Excel : TExcel;
  old_lang: lcid;
  row : integer;
  cell : string;
  //  переменные цикла
  i_snd : integer;  // цикл по поставщикам угля
  i_rcv : integer;  // цикл по получателям угля
  i_srt : integer;  // цикл по маркам угля

  // вспомагательные переменные
  CoalSender : TCoalSender;
  CoalReceiver : TCoalReceiver;
  CoalReceiverList : TList;
  CoalSort : TCoalSort;
  CoalSortList : TList;
  // переменная содержит данные для вывода одной строки отчета
  CoalMessage : TCoalMessage;
const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
begin
  old_lang := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  try
    Excel.AddWorkBook(dest_filename);
    Excel.Visible := true;
  except
    raise Exception.Create('Невозможно загрузить Excel');
  end;
  try

    row := idest_rowOfData;

    // начало цикла обхода всего списка предприятий-поставщиков угля
    for i_snd :=1 to CoalSenderList.Count do begin
      CoalSender := CoalSenderList.Items[i_snd-1];
      CoalSender.GetInfo(CoalReceiverList);
      CoalMessage.coal_sender := CoalSender.coal_sender;

      // начало цикла обхода всего списка получателей угля по одному поставщику
      for i_rcv :=1 to CoalReceiverList.Count do begin
        CoalReceiver := CoalReceiverList.Items[i_rcv-1];
        CoalReceiver.GetInfo(CoalSortList);
        CoalMessage.cargo_receiver := CoalReceiver.cargo_receiver;
        CoalMessage.dest_coal := CoalReceiver.dest_coal;

        // начало цикла обхода всего списка марок угля по одному получателю
        for i_srt :=1 to CoalSortList.Count do begin
          CoalSort := CoalSortList.Items[i_srt-1];
          CoalSort.GetInfo(CoalMessage);
          //---------------------------------------------
          // вывод данных в Excel
          //---------------------------------------------
          cell := sdest_colSender + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.coal_sender; // поставщик угля
          //
          cell := sdest_colCoalSort + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.sort_name; // марка угля
          //
          cell := sdest_colCoalQnty + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.qnty; // кол-во угля
          //
          cell := sdest_colSum_free_nds + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.amount_free_nds +
                              CoalMessage.avg_pripl_skidki_free_nds; // сумма со средними приплатами/скидками
          //
          cell := sdest_colCoalReceiver + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.cargo_receiver; // грузополучатель угля
          //
          cell := sdest_colNotes + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.dest_coal; // назначение угля

          row := row + 1;
          //
        end;  // конец цикла обхода всего списка марок угля по одному получателю
        //
        cell := sdest_colCoalSort + IntToStr(row);
        Excel.Cell[cell] := sdest_nameItogo; // вывод строки названия итогов
        //
        row := row + 2;
        //
      end;  // конец цикла обхода всего списка получателей угля по одному поставщику
      //
      row := row + 4;
      //
    end;  // конец цикла обхода всего списка предприятий-поставщиков угля

  finally
    Excel.free;
    SetThreadLocale(old_lang);
  end;
end;
//---------------------------------------
destructor TCoal.Destroy;
Var
  i : integer;
  CoalSender : TCoalSender;
begin
  // обход всего списка и удаление всех элементов списка
  // из памяти
  for i :=1 to CoalSenderList.Count do begin
    CoalSender :=  CoalSenderList.Items[i-1];
    CoalSender.Free;
  end;
  inherited Destroy;
end;
//---------------------------------------
procedure TCoal.Calc;
Var
  i : integer;
  CoalSender : TCoalSender;
begin
  // обход всего списка и вызов метода Calc
  for i :=1 to CoalSenderList.Count do begin
    CoalSender :=  CoalSenderList.Items[i-1];
    CoalSender.Calc;
  end;
end;
//---------------------------------------
procedure TCoal.Read(src_filename : string);
Var
  Excel : TExcel;
  old_lang: lcid;
  CoalMessage : TCoalMessage;
  row : integer;
  end_of_file : boolean;
  cell : string;
const
  English_LOCALE = (LANG_ENGLISH + SUBLANG_DEFAULT * 1024) + (SORT_DEFAULT shl 16);
begin
  old_lang := GetThreadLocale;
  SetThreadLocale(English_Locale);

  Excel := TExcel.Create;
  try
    Excel.AddWorkBook(src_filename);
//    Excel.SelectWorkSheet(sSheetNameOfsrcData);
//    Excel.Visible := true;
  except
    raise Exception.Create('Невозможно загрузить Excel');
  end;
  try
    row := isrc_rowOfData;
    end_of_file := false;

    while not end_of_file do begin
      cell := 'A' + IntToStr(row);
      CoalMessage.dest_coal := Excel.Cell[cell];    // выбираем  назначение угля
      //
      cell := 'B' + IntToStr(row);
      CoalMessage.coal_sender := Excel.Cell[cell];  // выбираем  поставщика угля
      //
      cell := 'C' + IntToStr(row);
      CoalMessage.sort_name := Excel.Cell[cell];    // выбираем  марку угля
      //
      cell := 'D' + IntToStr(row);
      CoalMessage.qnty := Excel.Cell[cell];          // выбираем  количество угля
      //
      cell := 'E' + IntToStr(row);
      CoalMessage.amount_free_nds := Excel.Cell[cell];  // выбираем  сумму угля без НДС
      //
      cell := 'F' + IntToStr(row);
      CoalMessage.amount_whith_nds := Excel.Cell[cell]; // выбираем  сумму угля с НДС
      //
      CoalMessage.nds := CoalMessage.amount_free_nds/5; // считаем  НДС по углю
      //
      cell := 'G' + IntToStr(row);
      CoalMessage.cargo_receiver := Excel.Cell[cell];   // выбираем  грузополучателя
      //
      CoalMessage.readFlag := false;                    // устанавливаем флаг обработки сообщения в false

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
      AddToSelf(CoalMessage);
      //
      row := row + 1;
    end;

  finally
    Excel.free;
    SetThreadLocale(old_lang);
  end;
end;
//---------------------------------------
//  процедура создания отчета по углям для тряпки
//---------------------------------------
procedure TCoalReportForm.CreatCoalReport;
begin
  try
    Coal := TCoal.Create;
    Coal.Read(coalOpenDialog.FileName);
    Coal.Calc;
    Coal.Report(PathToProgram + '\Template\' + sdest_filename);
  finally
    Coal.Free;
  end;
end;

//  обработчик нажатия кнопки на панели инструментов
procedure TcoalReportForm.sbReportToExcelClick(Sender: TObject);
begin

  case mainPageControl.ActivePage.TabIndex of

    ispacingCoalTabSheet :
       begin
         CreatCoalReport;
       end; // конец ispacingCoalTabSheet
  end;  // end of CASE


//  ReportHeader := ReportHeader + 'НДС за период с ' +
//                  ndsBeginMaskEdit.Text + ' по ' + ndsEndMaskEdit.Text ;                  ' ' + '(' + name  + ')';

  Application.BringToFront;
  MessageDlg('Экспорт в Excel завершен', mtInformation, [mbOk], 0);
end;

procedure TcoalReportForm.ExitSpeedButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TcoalReportForm.OpenButtonClick(Sender: TObject);
begin
  if coalOpenDialog.Execute then begin
    sbReportToExcel.Enabled := true;
    fileNameLabel.Caption := coalOpenDialog.FileName;
  end;
end;


end.
