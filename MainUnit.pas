unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Menus, DBTables, main_type, ActnList,
  ComCtrls, Db, Buttons, ToolWin, ImgList, ExtCtrls, xDBTree;

type
  TDISMainForm = class(TForm)
    dsTV: TDataSource;
    tbTV: TTable;
    mainStatusBar: TStatusBar;
    mainCoolBar: TCoolBar;
    mainToolBar: TToolBar;
    sbShowForm: TSpeedButton;
    SpeedButton3: TSpeedButton;
    mainImageList: TImageList;
    tbDLL: TTable;
    dbTV: TXDBTreeView;
    disSession: TSession;
    ToolButton1: TToolButton;
    prikolTimer: TTimer;
    procedure ExitMenuItemClick(Sender: TObject);
    procedure AboutMenuItemClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure dbTVFillNode(Sender: TObject; Node: TTreeNode);
    procedure sbShowFormClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure dbTVDeletion(Sender: TObject; Node: TTreeNode);
    procedure SpeedButton3Click(Sender: TObject);
    procedure prikolTimerTimer(Sender: TObject);
  private
    { Private declarations }
  public                           
    { Public declarations }
    Config : TConfig;
  end;

var
  DISMainForm: TDISMainForm;
  i1 : integer;
  my_s :string;
  my_p : pointer;

implementation

uses  ZastUnit, MainDataUnit, shared_type, prikol_form;

{$R *.DFM}

procedure TDISMainForm.ExitMenuItemClick(Sender: TObject);
begin
  DISMainForm.Close;
end;

procedure TDISMainForm.AboutMenuItemClick(Sender: TObject);
begin
//  ZastForm.ShowModal;
end;

procedure TDISMainForm.FormCreate(Sender: TObject);
Var
  i : integer;
  temp : string;
  //
  username : array[0..50] of char;
  p_username : pchar;
  len_name : cardinal;
  //
  Len : integer;
begin
  // пытаемс€ получить в качестве параметра
  // им€ файла конфигурации отчетной системы
  try
    prikolTimer.Enabled := false;
    // переключаем раскладку клавиатуры на русский
    LoadKeyboardLayout('00000419', KLF_ACTIVATE); // русский
    // LoadKeyboardLayout('00000409', KLF_ACTIVATE); // английский

    Config := TConfig.Create;
    if (MainDataModule = nil) then
      raise Exception.Create('Ќе инициализированы базы данных');
    // копируем информацию о приложении в TConfig
    Config.conf.DB := MainDataModule.DatabaseDIS;
    Config.conf.DBcyrr := MainDataModule.DatabaseDIScyrr;
    Config.conf.ora_isdDB := MainDataModule.DatabaseISD2000;
    Config.conf.dis_ibdb := MainDataModule.Database_dis_ibdb_cyrr;
    // получаем первый переданный параметр из командной строки
    Config.conf.nameProfile := ParamStr(1);
    if (Config.conf.nameProfile = '') then
      raise Exception.Create('Ќе задано им€ файла профил€ в командной строке');

    // получаем путь к каталогу из которого был запущен exe-файл.
    Config.conf.PathToProgram := Application.ExeName;
    temp := Config.conf.PathToProgram;
    Len := Length(Config.conf.PathToProgram);
    for i:=Len downto 1 do begin
       if Config.conf.PathToProgram[i]='\' then begin
         Config.conf.PathToProgram[i] := ' ';
         Break;
       end;
       Config.conf.PathToProgram[i] := ' ';
    end;
    Config.conf.PathToProgram := TrimRight(Config.conf.PathToProgram);
    disSession.NetFileDir := Config.conf.PathToProgram;
    Config.dbTV := dbTV;
    Config.conf.tbProfile := tbDLL;
    Config.conf.Owner := Application;
    Config.conf.nameProfile := Config.conf.PathToProgram + '\' + Config.conf.nameProfile;
    tbTV.TableName := Config.conf.nameProfile;
    // получаем им€ пользовател€ в системе
    p_username := @username;
    len_name := 30;
    GetUserName(p_username,cardinal(len_name));
    Config.conf.username := p_username;
    // проверка на права дл€ использовани€
    // файла конфигурации с ограниченным доступом
    //  yvoleynik, gural2, dtkachenko
    if ((ParamStr(1) = configPrivateAccess) and
        ((Config.conf.username <> 'gural2') and
         (Config.conf.username <> 'yvoleynik') and
         (Config.conf.username <> 'borzilo') and
         (Config.conf.username <> 'dtkachenko'))
       ) then
       raise Exception.Create('” вас нет прав дл€ работы с данной конфигурацией программы');
    //
    Config.Loaded := true;
  except
    raise;
  end;
end;                                     

procedure TDISMainForm.FormDestroy(Sender: TObject);
begin
  if Config <> nil then begin
    Config.Free;
    Config := nil;
  end;
end;

procedure TDISMainForm.dbTVFillNode(Sender: TObject; Node: TTreeNode);
Var
  ID : integer;
  Link : TLink;
  FormDLL : TReportFormDLL;
  NodeParam : TNodeParam;
  NodeInfo : shared_type.TNodeInfo;
begin
  ID := (Node as TXTreeNode).ID;
  Link := Config.conf.LinkList.GetLink(1,ID);
  // Link.side_2 - идентификатор библиотеки DLL
  FormDLL := Config.conf.LibDLL.GetDLL(Link.side_2);
  if FormDLL <> nil then begin
    NodeInfo := Config.conf.Profile.GetNode(ID);
    // записываем передаваемый в DLL параметр
    NodeParam := TNodeParam.Create;
    NodeParam.param := NodeInfo.param;
    NodeParam.pShowForm := FormDLL.ShowForm;
    (Node as TXTreeNode).Data := Pointer(NodeParam);
  end;
end;

procedure TDISMainForm.sbShowFormClick(Sender: TObject);
Var
  NodeParam : TNodeParam;
  pShowForm : ptrShowFormObject;
begin
  NodeParam := TNodeParam((dbTV.Selected as TXTreeNode).Data);
  if NodeParam <> nil then begin
    pShowForm := NodeParam.pShowForm;
//    if pShowForm <> nil then
      pShowForm(NodeParam.param);
  end;
end;

procedure TDISMainForm.dbTVDeletion(Sender: TObject; Node: TTreeNode);
Var
  NodeParam : TNodeParam;
begin
  if (Node as TXTreeNode).Data <> nil then begin
    NodeParam := (Node as TXTreeNode).Data;
    NodeParam.Free;
  end;
end;

procedure TDISMainForm.SpeedButton3Click(Sender: TObject);
begin
  Close;
end;

procedure TDISMainForm.prikolTimerTimer(Sender: TObject);
begin
  if tag = 1 then begin
    tag := 0;
    exit;
  end;
  if prikolForm.canJoke and not prikolForm.keepShow then begin
    prikolForm.Show;
  end;
end;

end.
