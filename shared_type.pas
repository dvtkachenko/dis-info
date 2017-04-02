unit shared_type;
//  данный модуль содержит описание типов , которые
//  могут использоваться во всех модулях программного комплекса
//  в т.ч. в версиях Delphi 4, Delphi 5.

interface
uses SysUtils, Classes, Menus, Windows, Forms,
     DBTables, ComObj, OleCtrls;

const
  maxTreeItem = 40;
  maxPChar = 254;

  // режимы поиска продукции (для модуля service)
  iprod_mode = 1;
  iprod_type_mode = 2;
  iextra_item_mode = 3;

  // максимальное кол-во типов договоров
  // по которым может производиться вытяжка статистики
  // type_relation (для модуля contract)
  imaxCntrItem = 20;

  // имя файла конфигурации к которому надо ограничить доступ
  configPrivateAccess = 'analitic.db';

  // начальная дата в отчетах
  startDate = '01.01.2009';

  // маски для выделения из параметра передаваемого в ShowForm
  // доступных страниц
  // для кодирования унарным методом исп-ся два младших байта
  i1page = $00008000;    //  32768 - в дес. системе счисления
  i2page = $00004000;    //  16384 - в дес. системе счисления
  i3page = $00002000;    //   8192 - в дес. системе счисления
  i4page = $00001000;    //   4096 - в дес. системе счисления
  i5page = $00000800;    //   2048 - в дес. системе счисления
  i6page = $00000400;    //   1024 - в дес. системе счисления
  i7page = $00000200;    //    512 - в дес. системе счисления
  // маски для выделения из параметра передаваемого в ShowForm
  // доступных форм (максимум 8)
  // для кодирования унарным методом исп-ся третий старший байт
  // 4-старший байт зарезервирован
  i1form = $00010000;    //   65536 - в дес. системе счисления
  i2form = $00020000;    //  131072 - в дес. системе счисления
  i3form = $00040000;    //  262144 - в дес. системе счисления
  i4form = $00080000;    //  524288 - в дес. системе счисления

type

  TProfile = class;
  TLibDLL = class;
  TLinkList = class;
  TSharedDll = class;
  ptrShowFormObject = procedure(param : integer) of object;

  // для хранения информации о загруженных библиотеках
  // системной информации о программе
  // и текущей конфигурации
  _config = record
    tbProfile : TTable;
    nameProfile : string;   // полное имя файла конфигурации
    PathToProgram : string;
    DB : TDatabase;
    DBcyrr : TDatabase;
    ora_isdDB : TDatabase;
    dis_ibdb : TDatabase;
    Owner : TApplication;
    Profile : TProfile;
    LibDLL : TLibDLL;
    LinkList : TLinkList;
    SharedDll : TSharedDll;
    username : string;
  end;
  //
  p_config = ^_config;
  //
  TNodeInfo = record
    nodeID : integer;
    dllName : string;
    param : integer;
  end;

  TLink = record
    link_id : integer;
    side_1 : integer;
    side_2 : integer;
  end;

  // параметры вызова процедур обработки для каждого из узлов дерева
  TNodeParam = class
//    nodeID : integer;
    param : integer;
    pShowForm : ptrShowFormObject;
  end;

  TLinkList = class
  public
    countLink : integer;
    Links : array[1..maxTreeItem] of TLink;
    function AddLink(side_1,side_2 : integer) : boolean;
    function GetLink(side,id : integer): TLink;
  end;

  // класс для работы с DLL
  TDLL = class
  private
    FLoaded : boolean;
    FhLib : HINST;
    function GetLoaded : boolean;
    procedure SetLoaded(Value : boolean);
    function GethLib : HINST;
  protected
    procedure LoadDLL; virtual;
    procedure FreeDLL; virtual;
  public
    nameDLL : string; // путь и имя библиотеки
    property Loaded : boolean read GetLoaded write SetLoaded default false;
    property hLib : HINST read GethLib default 0;
    destructor Destroy; override;
  end;

  // класс для работы с DLL содержащих формы
  TReportFormDLL = class(TDLL)
  private
    FhProcInitDLL : Pointer;
    FhProcUnInitDLL : Pointer;
    FhProcShowForm : Pointer;
    FhProcInitServiceExternalCall : Pointer;
    FhProcDeInitServiceExternalCall : Pointer;
    FConf : p_config;
    procedure InitDLL;
    procedure UnInitDLL;
  public
    ID : integer;  // identifier
    constructor Create(Conf : p_config);
    procedure ShowForm(param : integer);
    procedure LoadDLL; override;
    procedure InitServiceExternalCall(var p : pointer);
    procedure DeInitServiceExternalCall;
    destructor Destroy; override;
  end;

  // класс для хранения считанной информации из файла конфигурации
  TProfile = class(TObject)
  private
  public
    NodesInfo : array[1..maxTreeItem] of TNodeInfo;
//    tbProfile : PTTable;
    countNode : integer;
    function Read(tbProfile : TTable) : boolean;
    function GetNode(nodeID : integer) : TNodeInfo;
  end;

  // класс для работы с множеством библиотек DLL содержащих формы
  TLibDLL = class
  private
    FLoaded : boolean;
    FConf : p_config;
    procedure LoadDlls;
    procedure FreeDlls;
    function GetLoaded : boolean;
    procedure SetLoaded(Value : boolean);
  public
    countDll : integer;
    iExecutingDLL : integer;
    PathToDll : string;
    DLLs : array[1..maxTreeItem] of TReportFormDLL;
    property Loaded : boolean read GetLoaded write SetLoaded default false;
    function GetDLL(ID : integer) : TReportFormDLL;
    function GetDLLbyFilename(filename : string) : TReportFormDLL;
    constructor Create(Conf : p_config);
    destructor Destroy; override;
  end;

  // класс для работы с DLL содержащая разделяемые ресурсы
  TSharedDLL = class(TDLL)
    private
    FhProcInitDLL : Pointer;
    FhProcUnInitDLL : Pointer;
    FhProcReadDate : Pointer;
    FhProcWriteDate : Pointer;
    procedure InitDLL;
    procedure UnInitDLL;
  public
    function ReadDate(Var _BeginDate, _EndDate : TDateTime):boolean;
    function WriteDate(_BeginDate, _EndDate : TDateTime):boolean;
    procedure LoadDLL; override;
    destructor Destroy; override;
  end;

  ptrGetMenuItem = procedure(menu_name : PChar);
  ptrInitDLL = procedure(conf : p_config);
  ptrUnInitDLL = procedure;
  ptrShowForm = procedure(param : integer);
  ptrInitServiceExternalCall = procedure(var p : pointer);
  ptrDeInitServiceExternalCall = procedure;
  //
  ptrSharedInitDLL = procedure;
  ptrProcReadDate = function(Var _BeginDate, _EndDate : TDateTime):boolean;
  ptrProcWriteDate = function(_BeginDate, _EndDate : TDateTime):boolean;

implementation

// ---------------------------------------
// --- реализация методов класса TLinkList
// ---------------------------------------
//---------------------------------------
function TLinkList.GetLink(side,id : integer) : TLink;
Var
  i : integer;
begin
//  Result := nil;
  Result.link_id := 0;
  Result.side_1 := 0;
  Result.side_2 := 0;
  for i := 1 to countLink do begin
    if (side = 1) then begin
      if (Links[i].side_1 = id) then begin
        Result.link_id := Links[i].link_id;
        Result.side_1 := Links[i].side_1;
        Result.side_2 := Links[i].side_2;
        exit;
      end;
    end;
    if (side = 2) then begin
      if (Links[i].side_2 = id) then begin
        Result.link_id := Links[i].link_id;
        Result.side_1 := Links[i].side_1;
        Result.side_2 := Links[i].side_2;
        exit;
      end;
    end;
  end;
end;
//---------------------------------------
function TLinkList.AddLink(side_1,side_2 : integer) : boolean;
begin
  Result := false;
  if ((countLink + 1) <= maxTreeItem) and ((countLink + 1) > 0)then begin
    countLink := countLink + 1;
    Links[countLink].link_id := countLink;
    Links[countLink].side_1 := side_1;
    Links[countLink].side_2 := side_2;
    Result := true;
  end;
end;

// ---------------------------------------
// --- реализация методов класса TDLL
// ---------------------------------------
procedure TDLL.LoadDLL;
Var
  pname : PChar;
  name : array[0..maxPChar] of Char;
begin
  if (FLoaded = false) and (FhLib < 32) then begin
    if nameDLL = '' then
      raise Exception.Create('Не задано имя библиотеки DLL');
    pname := @name;
    strPCopy(pname,nameDLL);
    FhLib := LoadLibrary(pname);
    if (FhLib < 32) then
      raise Exception.Create('Ошибка при загрузке DLL');
    FLoaded := true;
  end;
end;
//---------------------------------------
procedure TDLL.FreeDLL;
begin
  if (FhLib >= 32) and (FLoaded = true) then begin
    FreeLibrary(FhLib);
    FLoaded := false;
    FhLib := 0;
  end;
end;
//---------------------------------------
function TDLL.GetLoaded : boolean;
begin
  Result := FLoaded;
end;
//---------------------------------------
procedure TDLL.SetLoaded(Value : boolean);
begin
  if Value then
    LoadDLL
  else
    FreeDLL;
end;
//---------------------------------------
function TDLL.GethLib : HINST;
begin
  Result := FhLib;
end;
//---------------------------------------
destructor TDLL.Destroy;
begin
  Loaded := false;
  inherited Destroy;
end;

// ---------------------------------------
// --- реализация методов класса TReportFormDLL
// ---------------------------------------
//---------------------------------------
constructor TReportFormDLL.Create(Conf : p_config);
begin
  inherited Create;
  FConf := Conf;
end;
//---------------------------------------
destructor TReportFormDLL.Destroy;
begin
  UnInitDLL;
  FhProcInitDLL := nil;
  FhProcUnInitDLL := nil;
  FhProcShowForm := nil;
  inherited Destroy;
end;
//---------------------------------------
procedure TReportFormDLL.ShowForm(param : integer);
begin
  if FLoaded <> false then ptrShowForm(FhProcShowForm)(param);
end;
//---------------------------------------
procedure TReportFormDLL.InitDLL;
begin
  if FLoaded <> false then ptrInitDLL(FhProcInitDLL)(FConf);
end;
//---------------------------------------
procedure TReportFormDLL.UnInitDLL;
begin
  if FLoaded <> false then ptrUnInitDLL(FhProcUnInitDLL);
end;
//---------------------------------------
procedure TReportFormDLL.InitServiceExternalCall(var p : pointer);
begin
  if (FLoaded <> false) and (FhProcInitServiceExternalCall <> nil) then
    ptrInitServiceExternalCall(FhProcInitServiceExternalCall)(p)
  else
    if (FhProcInitServiceExternalCall = nil) then
      raise Exception.Create('Данный модуль DLL не поддерживает внешние вызовы');
end;
//---------------------------------------
procedure TReportFormDLL.DeInitServiceExternalCall;
begin
  if (FLoaded <> false) and (FhProcDeInitServiceExternalCall <> nil) then
    ptrDeInitServiceExternalCall(FhProcDeInitServiceExternalCall)
  else
    if (FhProcDeInitServiceExternalCall = nil) then
      raise Exception.Create('Данный модуль DLL не поддерживает внешние вызовы');
end;
//---------------------------------------
procedure TReportFormDLL.LoadDLL;
begin
  inherited LoadDLL;
  FhProcInitDLL := GetProcAddress(hLib,'InitDLL');
  FhProcUnInitDLL := GetProcAddress(hLib,'UnInitDLL');
  FhProcShowForm := GetProcAddress(hLib,'ShowForm');
  // необязательные процедуры . Результат может быть nil
  FhProcInitServiceExternalCall := GetProcAddress(hLib,'InitServiceExternalCall');
  FhProcDeInitServiceExternalCall := GetProcAddress(hLib,'DeInitServiceExternalCall');
  if (FhProcInitDLL = nil) or
     (FhProcUnInitDLL = nil) or
     (FhProcShowForm = nil) then
    raise Exception.Create('Ошибка при инициализации DLL');
  InitDLL;
end;
//---------------------------------------

// ---------------------------------------
// --- реализация методов класса TSharedDLL
// ---------------------------------------
destructor TSharedDLL.Destroy;
begin
  UnInitDLL;
  FhProcInitDLL := nil;
  FhProcUnInitDLL := nil;
  FhProcReadDate := nil;
  FhProcWriteDate := nil;
  inherited Destroy;
end;
//---------------------------------------
function TSharedDLL.ReadDate(Var _BeginDate, _EndDate : TDateTime):boolean;
begin
  if FLoaded <> false then
    ReadDate := ptrProcReadDate(FhProcReadDate)(_BeginDate, _EndDate)
  else
    ReadDate := false;
end;
//---------------------------------------
function TSharedDLL.WriteDate(_BeginDate, _EndDate : TDateTime):boolean;
begin
  if FLoaded <> false then
    WriteDate := ptrProcWriteDate(FhProcWriteDate)(_BeginDate, _EndDate)
  else
    WriteDate := false;
end;
//---------------------------------------
procedure TSharedDLL.InitDLL;
begin
  if FLoaded <> false then ptrSharedInitDLL(FhProcInitDLL);
end;
//---------------------------------------
procedure TSharedDLL.UnInitDLL;
begin
  if FLoaded <> false then ptrUnInitDLL(FhProcUnInitDLL);
end;
//---------------------------------------
procedure TSharedDLL.LoadDLL;
begin
  inherited LoadDLL;
  FhProcInitDLL := GetProcAddress(hLib,'InitDLL');
  FhProcUnInitDLL := GetProcAddress(hLib,'UnInitDLL');
  FhProcReadDate := GetProcAddress(hLib,'ReadDate');
  FhProcWriteDate := GetProcAddress(hLib,'WriteDate');
  if (FhProcInitDLL = nil) or
     (FhProcUnInitDLL = nil) or
     (FhProcReadDate = nil) or
     (FhProcWriteDate = nil) then
    raise Exception.Create('Ошибка при инициализации DLL');
  InitDLL;
end;

// ---------------------------------------
// --- реализация методов класса TLibDLL
// ---------------------------------------
//---------------------------------------
constructor TLibDLL.Create(Conf : p_config);
begin
  inherited Create;
  FConf := Conf;
end;
//---------------------------------------
destructor TLibDLL.Destroy;
begin
  Loaded := false;
  inherited Destroy;
end;
//---------------------------------------
procedure TLibDLL.LoadDLLs;
Var
  i : integer;
  limit : integer;
  name : string;
  tempDLL : TReportFormDLL;
begin
  if FLoaded <> true then begin
    countDLL := 0;
    if FConf.Profile.countNode > maxTreeItem then
      limit := maxTreeItem
    else
      limit := FConf.Profile.countNode;

    for i := 1 to limit do begin
      if (FConf.Profile.NodesInfo[i].dllName <> '') then begin
        name := FConf.Profile.NodesInfo[i].dllName;
        tempDLL := TReportFormDLL.Create(FConf);
        tempDLL.nameDLL := PathToDll + '\' + FConf.Profile.NodesInfo[i].dllName;
        tempDLL.Loaded := true;
        if tempDLL.Loaded <> false then
          begin
            countDll := countDll + 1;
            tempDll.ID := countDll;
            // заполняем информацию о связи библиотеки с
            // соответствующим узлом дерева
            // side_1 = nodeID
            // side_2 = DllID
            FConf.LinkList.AddLink(FConf.Profile.NodesInfo[i].nodeID,tempDll.ID);
//        tempDll.ID := FConfig.Profile.NodesInfo[i].nodeID;
            DLLs[countDll] := tempDLL;
          end
        else
          raise Exception.Create('Ошибка при инициализации модулей DLL');
      end;
    end;
    FLoaded := true;
  end;
end;
//---------------------------------------
procedure TLibDLL.FreeDLLs;
Var
  i : integer;
begin
  if FLoaded <> false then begin
    for i := 1 to countDLL do begin
      if DLLs[i] <> nil then begin
        DLLs[i].Free;
        DLLs[i] := nil;
      end;
    end;
    FLoaded := false;
  end;
end;
//---------------------------------------
function TLibDLL.GetDLL(ID : integer) : TReportFormDLL;
Var
  i : integer;
begin
  Result := nil;
  for i := 1 to countDll do begin
    if (DLLs[i].ID = ID) then begin
      Result := DLLs[i];
      exit;
    end;
  end;
end;
//---------------------------------------
function TLibDLL.GetDLLbyFilename(filename : string) : TReportFormDLL;
Var
  i : integer;
begin
  Result := nil;
  for i := 1 to countDll do begin
    if (ExtractFileName(DLLs[i].nameDLL) = filename) then begin
      Result := DLLs[i];
      exit;
    end;
  end;
end;
//---------------------------------------
function TLibDLL.GetLoaded;
begin
  Result := FLoaded;
end;
//---------------------------------------
procedure TLibDLL.SetLoaded(Value : boolean);
begin
  if Value then
    LoadDLLs
  else
    FreeDLLs;
end;

// ---------------------------------------
// --- реализация методов класса TProfile
// ---------------------------------------
// читает файл конфигурации
function TProfile.Read(tbProfile : TTable) : boolean;
Var
  S : string;
begin
  Read := false;
  countNode := 0;
  try
    tbProfile.Open;
    while not tbProfile.eof do begin
      countNode := countNode + 1;
      NodesInfo[countNode].nodeID := tbProfile.fieldbyname('id').asinteger;
      s := tbProfile.fieldbyname('NameDLL').asstring;
      if countNode > maxTreeItem then
        raise Exception.Create('Слишком много модулей dll');
      NodesInfo[countNode].dllName := trim(S);
      NodesInfo[countNode].param := tbProfile.fieldbyname('param').asinteger;
      tbProfile.Next;
    end;
    Read := true;
  finally
    tbProfile.Close;
  end;
end;
//---------------------------------------
function TProfile.GetNode(nodeID : integer) : TNodeInfo;
Var
  i : integer;
begin
  Result.nodeID := 0;
  Result.dllName := '';
  Result.param := 0;
  for i := 1 to countNode do begin
    if NodesInfo[i].nodeID = nodeID then begin
      Result.nodeID := NodesInfo[i].nodeID;
      Result.dllName := NodesInfo[i].dllName;
      Result.param := NodesInfo[i].param;
      exit;
    end;
  end;
end;

end.
