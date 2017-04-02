//----------------------------------------------------------------------------------------
// © VI, Elena Philippova , 2000
//----------------------------------------------------------------------------------------
unit ListUtils;

interface
Uses Classes , SysUtils;

Type
  TListsItem = class(TCollectionItem)
  Private
     FValue  : Variant;
     FName   : String;
     FData   : Pointer;
  Protected
     Function  GetAsInteger : LongInt;
     Procedure SetAsInteger(AValue : LongInt );

     Function  GetAsString : String;
     Procedure SetAsString(AValue : String );

     Function  GetAsCurrency : Currency;
     Procedure SetAsCurrency(AValue : Currency );

  Public
     procedure AssignTo( Dest: TPersistent ); override;
     property Value       : Variant  read FValue        write FValue;
     property Name        : String   read FName         write FName;
     property AsInteger   : LongInt  read GetAsInteger  write SetAsInteger;
     property AsString    : String   read GetAsString   write SetAsString;
     property AsCurrency  : Currency read GetAsCurrency write SetAsCurrency;
     property Data        : Pointer  read FData         write FData;

  End;

  TCollectionListItemClass = class (TListsItem);

  TLists    = class (TCollection)
  private
     function GetListItem(Index : Integer) : TListsItem;
  Public
     Constructor Create(ItemClass: TCollectionItemClass);
     Procedure   AssignTo(Dest: TPersistent); override;
     Procedure   MergeTo( Dest : TLists ) ;
     Procedure   AssignNamesTo(Dest: TPersistent);
     Function    AddItem( Value : Variant; AName :  String =''; AData : Pointer = nil ) : TListsItem; virtual;
     Procedure   FillFromArray(ArValue : array of Variant);
     Procedure   FillFromNamedArray(ArValue , ArName : array of Variant );

     Function    IndexOf( Value : Variant ) : Integer;
     Function    JoinList( Separator : String = ','; Prefix : String = '') : String;

     Function    GetFromName(AName : String ) : TListsItem;
     Function    GetValueFromName(AName : String; DefaultValue : Variant ) : Variant;

     Procedure   DeleteFromValue( Value : Variant; All : Boolean = FALSE);
     Procedure   DeleteFromName(AName : String );

     Procedure   InitItems( InitValue : Variant );

     Property    AnItems[Index : Integer] : TListsItem read GetListItem; default;
  End;

  Function EmptyToStr( Value : Integer; Empty : String = '') : String;
  Function MaxValue( Value1 , Value2 : Variant ) : Variant;

implementation

//----------------------------------------------------------------------------------------
Function EmptyToStr( Value : Integer; Empty : String = '') : String;
Begin
     IF Value = 0
     Then Result:=Empty
     Else Result:=IntToStr(Value);
End;
//----------------------------------------------------------------------------------------
Function MaxValue( Value1 , Value2 : Variant ) : Variant;
Begin

    IF Value1 > Value2
    Then Result:=Value1
    Else Result:=Value2;

End;
//----------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------
//                       TLists
//----------------------------------------------------------------------------------------
Constructor TLists.Create(ItemClass: TCollectionItemClass);
Begin
     Inherited Create(ItemClass);
End;
//----------------------------------------------------------------------------------------
function TLists.GetListItem(Index : Integer) : TListsItem;
Begin
    Result:=TListsItem(Items[Index]);
End;
//----------------------------------------------------------------------------------------
function TLists.AddItem(Value : Variant; AName : String = ''; AData : Pointer = nil ) : TListsItem;
Begin
    Result:=TListsItem(Self.Add);
    Result.FValue:=Value;
    Result.FName:=AName;
    Result.FData:=AData;
End;
//----------------------------------------------------------------------------------------
function TLists.IndexOf(Value : Variant): Integer;
begin
  Result := 0;
  while (Result < Count) and ( AnItems[Result].Value <> Value) do
    Inc(Result);
  IF Result = Count then  Result := -1;
end;
//----------------------------------------------------------------------------------------
Function TLists.JoinList( Separator : String = ','; Prefix : String = '') : String;
Var i : Integer;
Begin
   Result:='';

   IF Count > 0 Then
   Begin
     For i:=0 To Count-1 Do
      Result:= Result + Prefix + AnItems[i].AsString + Separator;

     Result:=Copy(Result , 1 , Length(Result)-1 );
   End;

End;
//----------------------------------------------------------------------------------------
Procedure TLists.DeleteFromValue( Value : Variant; All : Boolean = FALSE);
Var i : Integer;
Begin
   i:=IndexOf(Value);
   IF i >= 0 Then Delete(i);
End;
//----------------------------------------------------------------------------------------
Procedure  TLists.DeleteFromName(AName : String );
Var AItem : TListsItem;
Begin
   AItem:=GetFromName(AName);

   IF AItem <> nil Then Delete(AItem.Index);

End;
//----------------------------------------------------------------------------------------
Function  TLists.GetFromName(AName : String ) : TListsItem;
Var i : Integer;
Begin
   Result:=nil;

   For i:=0 To Count-1 Do
   IF CompareText(AnItems[i].FName , AName) = 0
   Then Begin
            Result:=AnItems[i];
            Exit;
        End;

End;
//----------------------------------------------------------------------------------------
Function TLists.GetValueFromName(AName : String; DefaultValue : Variant ) : Variant;
Begin
    Result:=DefaultValue;

    IF GetFromName(AName) <> nil Then Result:= GetFromName(AName).Value;
End;
//----------------------------------------------------------------------------------------
Procedure TLists.FillFromArray(ArValue : array of Variant);
Var i : Integer;
Begin
    Clear;

    For  i:=Low(ArValue) TO High(ArValue) Do AddItem(ArValue[i]);
End;
//----------------------------------------------------------------------------------------
Procedure TLists.FillFromNamedArray(ArValue , ArName : array of Variant );
Var i , No : Integer;
Begin
    FillFromArray(ArValue);

    No:=High(ArName);
    IF No > High(ArValue) Then No:=High(ArValue);

    For  i:=Low(ArName) TO No Do AnItems[i].FName:=ArName[i] ;
End;
//----------------------------------------------------------------------------------------
Procedure   TLists.InitItems( InitValue : Variant );
Var i : Integer;
Begin
     For i:=0 To Count-1 Do
     AnItems[i].Value:=InitValue;
End;
//----------------------------------------------------------------------------------------
Procedure   TLists.AssignTo( Dest: TPersistent);
Var i : Integer;
Begin
    IF Dest Is TStrings
    Then Begin

             TStrings(Dest).Clear;

             For i:=0 To Count-1 Do
             TStrings(Dest).AddObject( AnItems[i].AsString , AnItems[i].FData );

             Exit;
         End;

    IF Dest Is TLists
    Then  Begin

              TLists(Dest).Clear;

              For i:=0 To Count-1 Do
              AnItems[i].AssignTo( TLists(Dest).Add );

              Exit;
          End;

    Inherited AssignTo(Dest);
End;
//----------------------------------------------------------------------------------------
Procedure TLists.MergeTo( Dest : TLists ) ;
Var i : Integer;
Begin
   IF Dest = nil Then Exit;

   For i:=0 To Count-1 Do
   IF Dest.IndexOf(AnItems[i].Value) < 0 Then Dest.AddItem(AnItems[i].Value);

End;
//----------------------------------------------------------------------------------------
Procedure   TLists.AssignNamesTo(Dest: TPersistent);
Var i : Integer;
Begin
    IF Dest Is TStrings
    Then Begin

             TStrings(Dest).Clear;

             For i:=0 To Count-1 Do
              TStrings(Dest).AddObject( AnItems[i].Name , AnItems[i].FData );

         End;
End;
//----------------------------------------------------------------------------------------


//****************************************************************************************

//----------------------------------------------------------------------------------------
//                       TListItem
//----------------------------------------------------------------------------------------
procedure TListsItem.AssignTo( Dest: TPersistent );
Begin

    IF Dest Is TListsItem Then
    Begin
       TListsItem(Dest).FValue:=FValue;
       TListsItem(Dest).FName:=FName;
       TListsItem(Dest).FData:=FData;
    End
    Else inherited;
End;
//----------------------------------------------------------------------------------------
Function  TListsItem.GetAsInteger : LongInt;
Begin
      if TVarData(FValue).VType <> varNull then Result := FValue else Result := 0;
End;
//----------------------------------------------------------------------------------------
Procedure TListsItem.SetAsInteger(AValue : LongInt );
Begin
    FValue:=AValue;
End;
//----------------------------------------------------------------------------------------
Function  TListsItem.GetAsString : String;
Begin
    Result:=VarToStr(FValue);
End;
//----------------------------------------------------------------------------------------
Procedure TListsItem.SetAsString(AValue : String );
Begin
    FValue:=AValue;
End;
//----------------------------------------------------------------------------------------
Function  TListsItem.GetAsCurrency : Currency;
Begin
      if TVarData(FValue).VType <> varNull then Result := FValue  else Result := 0;
End;
//----------------------------------------------------------------------------------------
Procedure TListsItem.SetAsCurrency(AValue : Currency );
Begin
    FValue:=AValue;
End;
//----------------------------------------------------------------------------------------


end.
