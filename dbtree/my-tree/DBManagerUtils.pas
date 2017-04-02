unit DBManagerUtils;

interface
Uses DBtables , Classes , DB;

Procedure SetQuerySQL(const Strings: array of String; Query: TQuery);
Function GetQuery(const Strings: array of String; AOwner: TComponent): TQuery;
Function GetTable(const ATableName: String; AOwner: TComponent): TTable;
Procedure SetDataBaseName( Name : String );

implementation
Var
   __DataBaseName : String;
//------------------------------------------------------------------------------
Procedure SetDataBaseName( Name : String );
Begin
   __DataBaseName:=Name;
End;
//------------------------------------------------------------------------------
Procedure SetQuerySQL(const Strings: array of String; Query: TQuery);
var
  i: Integer;
begin
  with Query do begin
    SQL.Clear;
    for i := 0 to High(Strings)
      do SQL.Add(Strings[i]);
    for i := 0 to ParamCount-1 do Params[i].DataType := ftString;
  end;
end;
//------------------------------------------------------------------------------
Function GetQuery(const Strings: array of String; AOwner: TComponent): TQuery;
begin
  Result := TQuery.Create(AOwner);
  with Result do
  Begin
       DataBaseName:=__DataBaseName;
  End;

  SetQuerySQL(Strings, Result);

end;
//------------------------------------------------------------------------------
Function GetTable(const ATableName: String; AOwner: TComponent): TTable;
begin
  Result := TTable.Create(AOwner);
  with Result do
  Try
    DataBaseName:=__DataBaseName;
    TableName := ATableName;
    UpdateMode := upWhereKeyOnly;
    Open;
  Except
  end;
end;
//------------------------------------------------------------------------------
end.
