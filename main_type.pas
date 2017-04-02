unit main_type;

interface
uses SysUtils, Classes, Menus, Windows, Forms,
     DBTables, xDBTree, Excel_TLB, ComObj, OleCtrls, shared_type;

const
  sNameSharedDll = 'shared.dll';

type
  // класс для хранения информации о загруженных библиотеках
  // системной информации о программе
  // и текущей конфигурации
  TConfig = class(TObject)
  private
    FLoaded : boolean;
    function GetLoaded : boolean;
    procedure SetLoaded(Value : boolean);
    procedure LoadConfig;
    procedure FreeConfig;
  public
    dbTV : TXDBTreeView;
    conf : _config;
    //
    property Loaded : boolean read GetLoaded write SetLoaded default false;
    constructor Create;
    destructor Destroy; override;
  end;

implementation


// ---------------------------------------
// --- реализация методов класса TConfig
// ---------------------------------------
//---------------------------------------
constructor TConfig.Create;
begin
  inherited Create;
  conf.LibDLL := TLibDLL.Create(@conf);
  conf.Profile := TProfile.Create;
  conf.LinkList := TLinkList.Create;
  conf.SharedDll := TSharedDll.Create;
end;
//---------------------------------------
destructor TConfig.Destroy;
begin
  if Loaded <> false then
    Loaded := false;
  conf.LibDLL.Free;
  conf.Profile.Free;
  conf.LinkList.Free;
  conf.SharedDll.Free;
  inherited Destroy;
end;

// --------------------------------------
function TConfig.GetLoaded : boolean;
begin
  Result := FLoaded;
end;
// --------------------------------------
procedure TConfig.SetLoaded(Value : boolean);
begin
  if Value then
    LoadConfig
  else
    FreeConfig;
end;
// --------------------------------------
procedure TConfig.LoadConfig;
begin
  if (FLoaded <> true) and (conf.tbProfile <> nil) and (dbTV <> nil)  then begin
    conf.tbProfile.TableName := conf.nameProfile;
    conf.Profile.Read(conf.tbProfile);
    conf.LibDLL.PathToDll := conf.PathToProgram + '\report_dll';
    conf.LibDLL.Loaded := true;
    dbTV.DataSource.DataSet.Open;
    conf.SharedDll.nameDLL := conf.PathToProgram + '\' + sNameSharedDll;
    conf.SharedDll.LoadDLL;
    FLoaded := true;
  end;
end;
// --------------------------------------
procedure TConfig.FreeConfig;
begin
  if FLoaded <> false then begin
    conf.LibDLL.Loaded := false;
    dbTV.DataSource.DataSet.Close;
    FLoaded := false;
  end;
end;

end.
