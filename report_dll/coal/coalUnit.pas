unit coalUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Grids, DBGrids, ComObj, Mask, Buttons, ExtCtrls, Db,
  DBTables, ToolWin;

const
  ispacingCoalTabSheet = 0;
  sPriplSkidki = '';   // �� ���� ������ ��������� TCoalMessage.sort_name
                       // ��� ���� ����� ����� ����������� ���������������
                       // �� ������ � ������  TCoalReceiver
  sdest_filename = 'coal_for_balans.xlt'; // ��� ����� ������
  sSheetNameOfsrcData = 'coal';  // ��� ����� � ��������� �������

  idest_rowOfData = 2;       // ������ ������ � ����� ������
  sdest_colSender = 'A';
  sdest_colCoalSort = 'C';
  sdest_colCoalQnty = 'D';
  sdest_colSum_free_nds = 'E';
  sdest_colCoalReceiver = 'G';
  sdest_colNotes = 'H';
  sdest_nameItogo = '�����:';  // ������������ ������ ������
                               // ��� ������ � �����

  isrc_rowOfData = 2;       // ������ ������ � �������� �����
  sEndOfData = '';     // ���� Excel-������ ����� ����� ��������
                       // �� ������  �������� ����� ������ � �����

//  sndsReportTemplate = 'nds_report.xlt';
//  iMaxDept = 10;

type
  // �������� ���������� �� ������ Excel
  TCoalMessage = record
    dest_coal : string;       // ���������� �����
    coal_sender : string;     // ��������� �����
    sort_name : string;       // ����� ����
    qnty : real;
    amount_free_nds : real;
    amount_whith_nds : real;
    nds : real;
    cargo_receiver : string;  // ���������������
    // ��������������� ����������
    avg_pripl_skidki_free_nds : real; // ���������� ���������� ������������ ����������������������
                                      // �������� �������/������
    readFlag : boolean;       // ��������������� � true ���� ���������
                              // �������� � ����������
  end;

  // �������� ���� ���������� � ����� ���� ��� ���������� � �����
  TCoalSort = class
  private
    sort_name : string;
    qnty : real;
    amount_free_nds : real;
    amount_whith_nds : real;
    nds : real;
    // ���������� ���������� ����������������������
    // �������� �������/������
    avg_pripl_skidki_free_nds : real;
  public
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
    procedure GetInfo(Var CoalMessage : TCoalMessage);
    procedure Calc(avg:real); // ������������� ���������������������
    constructor Create(Var CoalMessage : TCoalMessage);
    destructor Destroy; override;
  end;

  //
  TCoalReceiver = class
  private
    CoalSortList : TList;
  public
    dest_coal : string;       // ���������� �����
    cargo_receiver : string;  // ���������������
    pripl_skidki_free_nds : real;      // ���������� ���������� ��������
                              // ������� ���������� ������������
                              // ��������������������� ���-�� �� ��� ��-��
                              // ������������ � CoalSortList
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
    procedure GetInfo(Var _CoalSortList : TList);
    procedure Calc; // ������������� ���������������������
    constructor Create(Var CoalMessage : TCoalMessage);
    destructor Destroy; override;
  end;

  // ���� ����� �� ����������
  TCoalSender = class
  private
    CoalReceiverList : TList;
  public
    coal_sender : string;     // ��������� �����
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
    procedure GetInfo(Var _CoalReceiverList : TList);
    procedure Calc; // ������������� ���������������������
    constructor Create(Var CoalMessage : TCoalMessage);
    destructor Destroy; override;
  end;

  // ���� �����
  TCoal = class
  private
    CoalSenderList : TList;
    procedure AddToSelf(Var CoalMessage : TCoalMessage);
  public
    procedure Read(src_filename : string); // ��������� ������ ������ �� Excel-�����
    procedure Calc; // ������������� ���������������������
    procedure Report(dest_filename : string); // ��������� ������������ ������ � Excel
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
    PathToProgram : string;   // ���� � ����� ���������
    Coal : TCoal;   //  �������� ����� �� ����� Excel � ����������������� ���� 
  end;

implementation

uses excel_type;

{$R *.DFM}

{��������� ���������}

{-------------------}

// ---------------------------------------
// --- ���������� ������� ������ TCoalSort
// ---------------------------------------
constructor TCoalSort.Create(Var CoalMessage : TCoalMessage);
begin
  inherited Create;
  // ������������� �������
  sort_name := CoalMessage.sort_name;
  qnty := CoalMessage.qnty;
  amount_free_nds := CoalMessage.amount_free_nds;
  amount_whith_nds := CoalMessage.amount_whith_nds;
  nds := CoalMessage.nds;
  avg_pripl_skidki_free_nds := 0;
  CoalMessage.readFlag := true; // ��������� ����� ��������� ���������
end;
//---------------------------------------
procedure TCoalSort.AddToSelf(Var CoalMessage : TCoalMessage);
begin
  // ���� ��������� ��� ���� ����� ��������� �� ������� �� ���������
  if (CoalMessage.readFlag) then exit;

  // ���� ����� ���� � ��������� ������������� �����
  // � ������ ������ �� ���������� ���������� ������ � ������ ������
  // �.�. �������� ��� ��������� ��������
  if (sort_name = CoalMessage.sort_name) then begin
    qnty := qnty + CoalMessage.qnty;
    amount_free_nds := amount_free_nds + CoalMessage.amount_free_nds;
    amount_whith_nds := amount_whith_nds + CoalMessage.amount_whith_nds;
    nds := nds + CoalMessage.nds;
    CoalMessage.readFlag := true; // ��������� ����� ��������� ���������
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
  // ��� ��������� ....
  inherited Destroy;
end;

// ---------------------------------------
// --- ���������� ������� ������ TCoalReceiver
// ---------------------------------------
constructor TCoalReceiver.Create(Var CoalMessage : TCoalMessage);
begin
  inherited Create;
  dest_coal := CoalMessage.dest_coal;       // ���������� �����
  cargo_receiver := CoalMessage.cargo_receiver;  // ���������������
  pripl_skidki_free_nds := 0;
  CoalSortList := TList.Create;
  // ���� � ������ ������ �������� �������� ������
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
  // ���� ��������� ��� ���� ����� ��������� �� ������� �� ���������
  if (CoalMessage.readFlag) then exit;
  // ���� ��������� ����� ������ ����������, �� ...
  if ((dest_coal = CoalMessage.dest_coal) and
      (cargo_receiver = CoalMessage.cargo_receiver)) then begin
     // ���� � ������ ���� �������� ������
     if (CoalMessage.sort_name = sPriplSkidki) then begin
       pripl_skidki_free_nds := pripl_skidki_free_nds + CoalMessage.amount_free_nds;
       CoalMessage.readFlag := true;
     end
     else begin
       // ���� ������ ��� ������
//       if (CoalSortList.Count = 0) then
//         CoalSortList.Add(TCoalSort.Create(CoalMessage));
       // ����� ����� ������ � ����� ���������� ���������
       for i :=1 to CoalSortList.Count do begin
         CoalSort := CoalSortList.Items[i-1];
         CoalSort.AddToSelf(CoalMessage);
         // ���� ��������� ��� ���� ��������� �� ������� �� �����
         if (CoalMessage.readFlag) then break;
       end;
       // ���������� ��� ������ �������� � ������ ����
       // ��������� �� ��������
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
  // ����� ����� ������ � ������� ������ ����������
  // ��������� ����������� � ������ ������
  for i :=1 to CoalSortList.Count do begin
    CoalSort := CoalSortList.Items[i-1];
    CoalSort.GetInfo(CoalMessage);
    all_qnty := all_qnty + CoalMessage.qnty;
  end;
  // ����� ����� ������ � ������ ����������������������
  // ������� �������/������
  for i :=1 to CoalSortList.Count do begin
    CoalSort := CoalSortList.Items[i-1];
    CoalSort.GetInfo(CoalMessage);
    avg := CoalMessage.qnty/all_qnty * pripl_skidki_free_nds;
    // ������� ��������/������
    CoalSort.Calc(avg);
  end;
end;
//---------------------------------------
destructor TCoalReceiver.Destroy;
Var
  i : integer;
  CoalSort : TCoalSort;
begin
  // ����� ����� ������ � �������� ���� ��������� ������
  // �� ������
  for i :=1 to CoalSortList.Count do begin
    CoalSort := CoalSortList.Items[i-1];
    CoalSort.Free;
  end;
  inherited Destroy;
end;

// ---------------------------------------
// --- ���������� ������� ������ TCoalSender
// ---------------------------------------
constructor TCoalSender.Create(Var CoalMessage : TCoalMessage);
begin
  inherited Create;
  coal_sender := CoalMessage.coal_sender;       // ��������� �����
  CoalReceiverList := TList.Create;
  CoalReceiverList.Add(TCoalReceiver.Create(CoalMessage));
end;
//---------------------------------------
procedure TCoalSender.AddToSelf(Var CoalMessage : TCoalMessage);
Var
  i : integer;
  CoalReceiver : TCoalReceiver;
begin
  // ���� ��������� ��� ���� ����� ��������� �� ������� �� ���������
  if (CoalMessage.readFlag) then exit;
  // ���� ��������� ����� ������ ����������, �� ...
  if (coal_sender = CoalMessage.coal_sender) then begin
    // ����� ����� ������ � ����� ���������� ���������
    for i :=1 to CoalReceiverList.Count do begin
      CoalReceiver := CoalReceiverList.Items[i-1];
      CoalReceiver.AddToSelf(CoalMessage);
      // ���� ��������� ��� ���� ��������� �� ������� �� �����
      if (CoalMessage.readFlag) then break;
    end;
    // ���������� ��� ������ �������� � ������ ����
    // ��������� �� ��������
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
  // ����� ����� ������ � ����� ������ Calc
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
  // ����� ����� ������ � �������� ���� ��������� ������
  // �� ������
  for i :=1 to CoalReceiverList.Count do begin
    CoalReceiver := CoalReceiverList.Items[i-1];
    CoalReceiver.Free;
  end;
  // ��� ��������� ....
  inherited Destroy;
end;

// ---------------------------------------
// --- ���������� ������� ������ TCoal
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
  // ���� ��������� ��� ���� ����� ��������� �� ������� �� ���������
  if (CoalMessage.readFlag) then exit;
  // ����� ����� ������ � ����� ���������� ���������
  for i :=1 to CoalSenderList.Count do begin
    CoalSender := CoalSenderList.Items[i-1];
    CoalSender.AddToSelf(CoalMessage);
    // ���� ��������� ��� ���� ��������� �� ������� �� �����
    if (CoalMessage.readFlag) then break;
  end;
  // ���������� ��� ������ �������� � ������ ����
  // ��������� �� ��������
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
  //  ���������� �����
  i_snd : integer;  // ���� �� ����������� ����
  i_rcv : integer;  // ���� �� ����������� ����
  i_srt : integer;  // ���� �� ������ ����

  // ��������������� ����������
  CoalSender : TCoalSender;
  CoalReceiver : TCoalReceiver;
  CoalReceiverList : TList;
  CoalSort : TCoalSort;
  CoalSortList : TList;
  // ���������� �������� ������ ��� ������ ����� ������ ������
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
    raise Exception.Create('���������� ��������� Excel');
  end;
  try

    row := idest_rowOfData;

    // ������ ����� ������ ����� ������ �����������-����������� ����
    for i_snd :=1 to CoalSenderList.Count do begin
      CoalSender := CoalSenderList.Items[i_snd-1];
      CoalSender.GetInfo(CoalReceiverList);
      CoalMessage.coal_sender := CoalSender.coal_sender;

      // ������ ����� ������ ����� ������ ����������� ���� �� ������ ����������
      for i_rcv :=1 to CoalReceiverList.Count do begin
        CoalReceiver := CoalReceiverList.Items[i_rcv-1];
        CoalReceiver.GetInfo(CoalSortList);
        CoalMessage.cargo_receiver := CoalReceiver.cargo_receiver;
        CoalMessage.dest_coal := CoalReceiver.dest_coal;

        // ������ ����� ������ ����� ������ ����� ���� �� ������ ����������
        for i_srt :=1 to CoalSortList.Count do begin
          CoalSort := CoalSortList.Items[i_srt-1];
          CoalSort.GetInfo(CoalMessage);
          //---------------------------------------------
          // ����� ������ � Excel
          //---------------------------------------------
          cell := sdest_colSender + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.coal_sender; // ��������� ����
          //
          cell := sdest_colCoalSort + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.sort_name; // ����� ����
          //
          cell := sdest_colCoalQnty + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.qnty; // ���-�� ����
          //
          cell := sdest_colSum_free_nds + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.amount_free_nds +
                              CoalMessage.avg_pripl_skidki_free_nds; // ����� �� �������� ����������/��������
          //
          cell := sdest_colCoalReceiver + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.cargo_receiver; // ��������������� ����
          //
          cell := sdest_colNotes + IntToStr(row);
          Excel.Cell[cell] := CoalMessage.dest_coal; // ���������� ����

          row := row + 1;
          //
        end;  // ����� ����� ������ ����� ������ ����� ���� �� ������ ����������
        //
        cell := sdest_colCoalSort + IntToStr(row);
        Excel.Cell[cell] := sdest_nameItogo; // ����� ������ �������� ������
        //
        row := row + 2;
        //
      end;  // ����� ����� ������ ����� ������ ����������� ���� �� ������ ����������
      //
      row := row + 4;
      //
    end;  // ����� ����� ������ ����� ������ �����������-����������� ����

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
  // ����� ����� ������ � �������� ���� ��������� ������
  // �� ������
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
  // ����� ����� ������ � ����� ������ Calc
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
    raise Exception.Create('���������� ��������� Excel');
  end;
  try
    row := isrc_rowOfData;
    end_of_file := false;

    while not end_of_file do begin
      cell := 'A' + IntToStr(row);
      CoalMessage.dest_coal := Excel.Cell[cell];    // ��������  ���������� ����
      //
      cell := 'B' + IntToStr(row);
      CoalMessage.coal_sender := Excel.Cell[cell];  // ��������  ���������� ����
      //
      cell := 'C' + IntToStr(row);
      CoalMessage.sort_name := Excel.Cell[cell];    // ��������  ����� ����
      //
      cell := 'D' + IntToStr(row);
      CoalMessage.qnty := Excel.Cell[cell];          // ��������  ���������� ����
      //
      cell := 'E' + IntToStr(row);
      CoalMessage.amount_free_nds := Excel.Cell[cell];  // ��������  ����� ���� ��� ���
      //
      cell := 'F' + IntToStr(row);
      CoalMessage.amount_whith_nds := Excel.Cell[cell]; // ��������  ����� ���� � ���
      //
      CoalMessage.nds := CoalMessage.amount_free_nds/5; // �������  ��� �� ����
      //
      cell := 'G' + IntToStr(row);
      CoalMessage.cargo_receiver := Excel.Cell[cell];   // ��������  ���������������
      //
      CoalMessage.readFlag := false;                    // ������������� ���� ��������� ��������� � false

      // ���� ������ �� ���������� � ����������, � ����������, � ���������������
      // �� ������ ����� ��� ��������� ����� ����� � �� ������� �� �����
      if ((CoalMessage.dest_coal = sEndOfData) and
          (CoalMessage.coal_sender = sEndOfData) and
          (CoalMessage.cargo_receiver = sEndOfData)) then begin
        end_of_file := true;
        continue;
      end;

      // ���� ������ �� ���������� ���� �� � ����� ��
      // ������������ �����, �� ���������� ���������� � �������
      // �� ���������
      if ((CoalMessage.dest_coal = sEndOfData) or
          (CoalMessage.coal_sender = sEndOfData) or
          (CoalMessage.amount_free_nds = 0) or
          (CoalMessage.cargo_receiver = sEndOfData)) then begin
        raise Exception.Create('����������� ������ � ����� �� ������������ �����');
      end;

      // �������� ��������� CoalMessage ������� Coal
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
//  ��������� �������� ������ �� ����� ��� ������
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

//  ���������� ������� ������ �� ������ ������������
procedure TcoalReportForm.sbReportToExcelClick(Sender: TObject);
begin

  case mainPageControl.ActivePage.TabIndex of

    ispacingCoalTabSheet :
       begin
         CreatCoalReport;
       end; // ����� ispacingCoalTabSheet
  end;  // end of CASE


//  ReportHeader := ReportHeader + '��� �� ������ � ' +
//                  ndsBeginMaskEdit.Text + ' �� ' + ndsEndMaskEdit.Text ;                  ' ' + '(' + name  + ')';

  Application.BringToFront;
  MessageDlg('������� � Excel ��������', mtInformation, [mbOk], 0);
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
