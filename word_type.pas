unit word_type;

interface
uses SysUtils, Classes, Menus, Windows, Forms,
     word_TLB, ComObj, OleCtrls;

type

  TWord = class
  private
    wa : _Application;
    wd : _Document;
    LCID : integer;
    procedure fSetVisible(Visible : boolean);
    function fGetVisible : boolean;
  public
    constructor Create;
    destructor Destroy; override;
    property Visible : boolean read fGetVisible write fSetVisible;
    procedure SelectDocument(name : string);
    procedure AddDocument(Template : OleVariant);
    procedure OpenDocument(FileName : OleVariant);
    procedure CloseDocument;
    procedure SaveAs(filename : string);
    // работа с документом
    procedure TypeText(text : WideString);
    procedure TypeParagraph;
    procedure EndKey;
  end;


implementation

// ---------------------------------------
// --- реализация методов класса TWord
// ---------------------------------------
constructor TWord.Create;
begin
  LCID := GetUserDefaultLCID;
  wa := CoApplication_.Create;
end;
//---------------------------------------
destructor TWord.Destroy;
Var
  SaveChs : OleVariant;
begin
  if Assigned(wa) then begin // а если он не создан?
    SaveChs := wdSaveChanges;
    wa.Quit(SaveChs,EmptyParam,EmptyParam);
    wa := nil;
  end;
  inherited;
end;
//---------------------------------------
procedure TWord.AddDocument(Template : OleVariant);
begin
  if Assigned(wa) then begin // а если он не создан?
    wd := wa.Documents.AddOld(Template, EmptyParam);
  end;
end;
//---------------------------------------
procedure TWord.OpenDocument(FileName : OleVariant);
begin
  if Assigned(wa) then begin // а если он не создан?
    wd := wa.Documents.OpenOld(
                              FileName,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam,
                              EmptyParam);
  end;
end;
//---------------------------------------
procedure TWord.SelectDocument(name : string);
Var
  wd : Word_TLB._Document;
  index : OleVariant;
begin
  if wa <> nil then begin
    index := name;
    wd := wa.Documents.Item(index) as Word_TLB._Document;
    OleVariant(wd).Select;
    wd := nil;
  end;
end;
//---------------------------------------
procedure TWord.fSetVisible(Visible : boolean);
begin
  if Assigned(wa) then begin // а если он не создан?
    wa.Visible := Visible;
    if Visible then begin
      if wa.WindowState = TOleEnum(wdWindowStateMinimize) then
        wa.WindowState := TOleEnum(wdWindowStateNormal);
      wa.ScreenUpdating := true;
    end;
  end;
end;
//---------------------------------------
function TWord.fGetVisible : boolean;
begin
  result := wa.Visible;
end;
//---------------------------------------
procedure TWord.CloseDocument;
Var
  SaveChs : OleVariant;
begin
  if Assigned(wa) then begin // а если он не создан?
    SaveChs := wdSaveChanges;
    wa.Documents.Close(SaveChs,EmptyParam,EmptyParam);
  end;
end;
//---------------------------------------
procedure TWord.SaveAs(filename : string);
Var
  _filename : OleVariant;
begin
  if Assigned(wa) then begin // а если он не создан?
    _filename := filename;
    wa.ActiveDocument.SaveAs(
        _filename,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam,
        EmptyParam);
  end;
end;
//---------------------------------------
procedure TWord.TypeText(text : WideString);
Var
  S: Selection;
begin
  S := wa.Selection;
  S.TypeText(text);
end;
//---------------------------------------
procedure TWord.TypeParagraph;
Var
  S: Selection;
begin
  S := wa.Selection;
  S.TypeParagraph;
end;
//---------------------------------------
procedure TWord.EndKey;
Var
  S: Selection;
  _wdStory : OleVariant;
begin
  S := wa.Selection;
  _wdStory := wdStory;
  S.EndKey(_wdStory,EmptyParam);
end;
//---------------------------------------
end.
