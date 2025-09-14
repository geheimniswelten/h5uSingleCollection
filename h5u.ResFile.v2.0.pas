/// <summary> Resource File Reader </summary>
/// <remarks> Version: 2.0 2025-09-14 <br/> Copyright 2025 himitsu @ geheimniswelten <br/> License: MPL v1.1 , GPL v3.0 or LGPL v3.0 </remarks>
/// <seealso cref="http://geheimniswelten.de"> Geheimniswelten </seealso>
/// <seealso cref="http://geheimniswelten.de/kontakt/#licenses"> License Text </seealso>
/// <seealso cref="https://github.com/geheimniswelten/h5uSingleCollection"> GitHub </seealso>
unit h5u.ResFile;

interface

uses
  System.RTLConsts, System.Math, System.Classes,
  System.SysUtils, System.StrUtils, System.IOUtils,
  System.Generics.Collections,
  Winapi.Windows;

type
  TResFile = record
  public type
    {$REGION 'SubTypes'}
    /// <summary> Resource Type Identifier, as String, PChar or Integer, including constants like RT_RCDATA </summary>
    /// <remarks> MAKEINTRESOURCE can also be used, but it would be easier to use the Integer directly. </remarks>
    TTypeOrID = record
    private
      FName: string;
      function  GetID:       Integer;
      procedure SetID(Value: Integer);
      function  GetRI:       PChar;
      procedure SetRI(Value: PChar);
      procedure SetName(const Value: string);
    public
      property Name:  string  read FName write SetName;
      property ID:    Integer read GetID write SetID;
      property ResID: PChar   read GetRI write SetRI;
      function isID:  Boolean;
      class operator Implicit(const Value: TTypeOrID): string;    inline;
      class operator Implicit(const Value: TTypeOrID): Integer;   inline;
      class operator Implicit(const Value: Integer):   TTypeOrID; inline;
      class operator Implicit(const Value: string):    TTypeOrID; inline;
      class operator Implicit(const Value: PChar):     TTypeOrID; inline;
    end;
    /// <summary> Resource Identifier as String or Integer </summary>
    /// <remarks> MAKEINTRESOURCE can also be used, but it would be easier to use the Integer directly. </remarks>
    TNameOrID = record
    private
      FName: string;
      function  GetID:       Integer;
      procedure SetID(Value: Integer);
      function  GetNI:       PChar;
      procedure SetName(const Value: string);
    public
      property Name: string  read FName write SetName;
      property ID:   Integer read GetID write SetID;
      function isID: Boolean;
      class operator Implicit(const Value: TNameOrID): string;    inline;
      class operator Implicit(const Value: TNameOrID): Integer;   inline;
      class operator Implicit(const Value: Integer):   TNameOrID; inline;
      class operator Implicit(const Value: string):    TNameOrID; inline;
      class operator Implicit(const Value: PChar):     TNameOrID; inline;
    end;
    /// <summary> Language ID, as in MAKELANGID, or as a string like "en-US" </summary>
    TLangID = record
    private
      FLangID: LANGID;
      function  GetLang(idx: Integer):       Word;
      procedure SetLang(idx: Integer; Value: Word);
      function  GetName:       string;
      procedure SetName(Value: string); inline;
    public
      property  LangID:          LANGID       read FLangID write FLangID;
      property  PrimaryLanguage: Word index 0 read GetLang write SetLang;
      property  SubLanguage:     Word index 1 read GetLang write SetLang;
      property  Language:        string       read GetName write SetName;
      procedure SetLangID(Primary, SubLang: Word); inline;
      class function GetDefaultLangID: TLangID; static; inline;
      class function GetLangID(LangName: string):    TLangID; static;
      class operator Implicit(const Value: TLangID): LANGID;  inline;
      class operator Implicit(const Value: LANGID):  TLangID; inline;
    end;
    TResVersion = record
      case Integer of
        0: ( Minor, Major: WORD  );
        1: ( Raw:          DWORD );
    end;
    TVersion = record
    private
      FData: TResVersion;
      function  GetVer:       Single;
      procedure SetVer(Value: Single);
      function  GetStr:       string; inline;
      procedure SetStr(Value: string);
    public
      property Minor:   WORD   read FData.Minor write FData.Minor;
      property Major:   WORD   read FData.Major write FData.Major;
      property Raw:     DWORD  read FData.Raw   write FData.Raw;
      property Version: Single read GetVer      write SetVer;
      property VerStr:  string read GetStr      write SetStr;
      class operator Implicit(const Value: TVersion): DWORD; inline;
      class operator Implicit(const Value: DWORD): TVersion; inline;
    end;
    TCharacteristics = DWORD;
    TResource = record
      Type_:           TTypeOrID;
      Name:            TNameOrID;
      ResVersion:      TResVersion;
      Flags:           Word;
      Language:        TLangID;
      Version:         TVersion;
      Characteristics: TCharacteristics;
      Data:            TBytes;
    private
      SortOrder:       Pointer;
      procedure ResolveDefaults;
      procedure ReadRecord (const FileData: TBytes; var Pos: Integer);
      procedure WriteRecord(var   FileData: TBytes; var Pos: Integer);
    public
      function DataSize: Integer; inline;
      function IsEmpty: Boolean;
      class operator Initialize(out Dest: TResource);
    end;
  private
    class function EnumResTypeProc(hModule: HMODULE; lpszType: PChar; lParam: NativeInt): BOOL; stdcall; static;
    class function EnumResNameProc(hModule: HMODULE; lpszType, lpszName: PChar; lParam: NativeInt): BOOL; stdcall; static;
    class function EnumResLangProc(hModule: HMODULE; lpszType, lpszName: PChar; wIDLanguage: Word; lParam: NativeInt): BOOL; stdcall; static;
    procedure SortEnumResources;
    {$ENDREGION}
  public class var
    DefaultLangID:  TLangID;   // 0 = neutral language / $FFFF = current system language (DEFAULT = $FFFF)
    DefaultVersion: TVersion;  // DEFAULT = 0
    DefaultFlags:   Word;      // DEFAULT = 0
  public
    Resources: TArray<TResource>;

    class function CheckIsPEFile(FileName: string): Boolean; static;
    class function Create: TResFile; overload; static; inline;
    constructor Create(FileName: string); overload;

    procedure Clear; inline;
    function  ResCount: Integer; inline;
    function  Add    (Resource: TResource): Integer; overload;
    function  Add    (&Type: TTypeOrID; Name: TNameOrID; Data: TBytes;                  Language: LANGID=$FFFF; Flags: Word=$FFFF): Integer; overload;
    function  IndexOf(&Type: TTypeOrID; Name: TNameOrID; IgnoreLanguage: Boolean=False; Language: LANGID=$FFFF):                    Integer; overload;
    function  IndexOf(Resource: TResource;               IgnoreLanguage: Boolean=False): Integer; overload;
    procedure Delete(idx: Integer); inline;

    function  Find(&Type: TTypeOrID):                  TArray<TResource>; overload;
    function  Find(&Type: TTypeOrID; Name: TNameOrID): TArray<TResource>; overload;
    function  FindName(              Name: TNameOrID): TArray<TResource>; overload;
    function  FindLang(Language: TLangID):             TArray<TResource>; overload;
    function  FindLanguages: TArray<TLangID>;  // DefaultLangID if empty

    /// <summary>
    ///   LoadMine loads the resources of its own EXE. <param/>
    ///   LoadFrom and SaveTo (Stream) loads and saves in *.res files. <param/>
    ///   And LoadFrom and SaveTo (FileName) processes *.res, *.exe, *.dll and *.bpl.
    /// </summary>
    procedure LoadMine;
    procedure LoadFrom(FileName: string); overload;
    procedure LoadFrom(FileData: TBytes); overload;
    procedure SaveTo  (FileName: string; RemoveMissingFromPEFiles: Boolean=False); overload;
    function  SaveTo: TBytes; overload;

    class constructor Create;
    class operator Initialize(out Dest: TResFile);
  public const
    {$REGION 'constants and helpers'}
    cDataTypes: array[#0..#24] of string = (
            '',
      {#1}  'RT_CURSOR',
      {#2}  'RT_BITMAP',
      {#3}  'RT_ICON',
      {#4}  'RT_MENU',
      {#5}  'RT_DIALOG',
      {#6}  'RT_STRING',
      {#7}  'RT_FONTDIR',
      {#8}  'RT_FONT',
      {#9}  'RT_ACCELERATOR',
      {#10} 'RT_RCDATA',
      {#11} 'RT_MESSAGETABLE',
      {#12} 'RT_GROUP_CURSOR',  // 11+RT_CURSOR
            '',
      {#14} 'RT_GROUP_ICON',    // 11+RT_ICON
            '',
      {#16} 'RT_VERSION',
      {#17} 'RT_DLGINCLUDE',
            '',
      {#19} 'RT_PLUGPLAY',
      {#20} 'RT_VXD',
      {#21} 'RT_ANICURSOR',
      {#22} 'RT_ANIICON',
      {#23} 'RT_HTML',
      {#24} 'RT_MANIFEST'
    );
    // the 4 languages of the Delphi IDE
    Lang_deDE    = LANGID($0407);  // MakeLangID(LANG_GERMAN,   SUBLANG_DEFAULT)
    Lang_enUS    = LANGID($0409);  // MakeLangID(LANG_ENGLISH,  SUBLANG_DEFAULT)
    Lang_frFR    = LANGID($040C);  // MakeLangID(LANG_FRENCH,   SUBLANG_DEFAULT)
    Lang_jaJP    = LANGID($0411);  // MakeLangID(LANG_JAPANESE, SUBLANG_DEFAULT)
    Lang_Neutral = LANGID($0000);  // MakeLangID(LANG_NEUTRAL,  SUBLANG_NEUTRAL)
  public
    /// <summary> e.q. RT_RCDATA to 'RT_RCDATA' </summary>
    class function ResTypeToStr(ResType: PChar): string; static;
    class function ResNameToStr(ResName: PChar): string; static;
    {$ENDREGION}
  end;

implementation

{ TResFile.TTypeOrID }

function TResFile.TTypeOrID.GetID: Integer;
begin
  if not StartsStr('#', FName) then begin
    Result := IndexText(FName, TResFile.cDataTypes);
    if (Result < 0) or (FName = '') then
      raise EReadError.CreateRes(@SInvalidString);
  end else
    Result := FName.Substring(1).ToInteger;
end;

function TResFile.TTypeOrID.GetRI: PChar;
begin
  if isID then
    Result := PChar(GetID)
  else
    Result := PChar(FName);
end;

class operator TResFile.TTypeOrID.Implicit(const Value: Integer): TTypeOrID;
begin
  Result.SetID(Value);
end;

class operator TResFile.TTypeOrID.Implicit(const Value: PChar): TTypeOrID;
begin
  Result.SetRI(Value);
end;

class operator TResFile.TTypeOrID.Implicit(const Value: string): TTypeOrID;
begin
  Result.SetName(Value);
end;

class operator TResFile.TTypeOrID.Implicit(const Value: TTypeOrID): Integer;
begin
  Result := Value.GetID;
end;

class operator TResFile.TTypeOrID.Implicit(const Value: TTypeOrID): string;
begin
  Result := Value.FName;
end;

function TResFile.TTypeOrID.isID: Boolean;
var
  Index: Integer;
begin
  if StartsStr('#', FName) then
    Result := Integer.TryParse(FName.Substring(1), Index)
  else
    Result := (FName <> '') and MatchText(FName, TResFile.cDataTypes);
end;

procedure TResFile.TTypeOrID.SetID(Value: Integer);
begin
  if (Value >= 0) and (Value <= Ord(High(TResFile.cDataTypes))) and (TResFile.cDataTypes[Char(Value)] <> '') then
    FName := TResFile.cDataTypes[Char(Value)]
  else
    FName := '#' + Value.ToString;
end;

procedure TResFile.TTypeOrID.SetName(const Value: string);
begin
  var Index := IndexText(Value, TResFile.cDataTypes);
  if (Index >= 0) or ( StartsStr('#', Value)
    and Integer.TryParse(FName.Substring(1), Index)
    and (Index >= 0) and (Index <= Ord(High(TResFile.cDataTypes)))
    and (TResFile.cDataTypes[Char(Index)] <> '') )
  then
    FName := TResFile.cDataTypes[Char(Index)]
  else
    FName := Value;
end;

procedure TResFile.TTypeOrID.SetRI(Value: PChar);
begin
  if NativeInt(Value) <= $0000_FFFF then
    SetID(NativeInt(Value))
  else
    SetName(Value);
end;

{ TResFile.TNameOrID }

function TResFile.TNameOrID.GetID: Integer;
begin
  if StartsStr('#', FName) then
    Result := FName.Substring(1).ToInteger
  else
    raise EReadError.CreateRes(@SInvalidString);
end;

function TResFile.TNameOrID.GetNI: PChar;
begin
  if isID then
    Result := PChar(GetID)
  else
    Result := PChar(FName);
end;

class operator TResFile.TNameOrID.Implicit(const Value: Integer): TNameOrID;
begin
  Result.SetID(Value);
end;

class operator TResFile.TNameOrID.Implicit(const Value: PChar): TNameOrID;
begin
  if NativeInt(Value) <= $0000_FFFF then
    Result.SetID(NativeInt(Value))
  else
    Result.SetName(Value);
end;

class operator TResFile.TNameOrID.Implicit(const Value: string): TNameOrID;
begin
  Result.SetName(Value);
end;

class operator TResFile.TNameOrID.Implicit(const Value: TNameOrID): Integer;
begin
  Result := Value.GetID;
end;

class operator TResFile.TNameOrID.Implicit(const Value: TNameOrID): string;
begin
  Result := Value.FName;
end;

function TResFile.TNameOrID.isID: Boolean;
var
  Index: Integer;
begin
  Result := StartsStr('#', FName) and Integer.TryParse(FName.Substring(1), Index);
end;

procedure TResFile.TNameOrID.SetID(Value: Integer);
begin
  FName := '#' + Value.ToString;
end;

procedure TResFile.TNameOrID.SetName(const Value: string);
begin
  FName := Value;
end;

{ TResFile.TLangID }

class function TResFile.TLangID.GetDefaultLangID: TLangID;
begin
  Result.FLangID := {GetUserDefaultLangID}GetThreadLocale;
end;

function TResFile.TLangID.GetLang(idx: Integer): Word;
begin
  case idx of
    0:   Result := Winapi.Windows.PRIMARYLANGID(FLangID);
    1:   Result := Winapi.Windows.SUBLANGID(FLangID);
    else Result := 0;
  end;
end;

class function TResFile.TLangID.GetLangID(LangName: string): TLangID;
begin
  Result.FLangID := LocaleNameToLCID(PChar(LangName), LOCALE_ALLOW_NEUTRAL_NAMES);
  if Result.FLangID = 0 then
    RaiseLastOSError;
  if Result.FLangID = LOCALE_CUSTOM_DEFAULT then
    Result := GetDefaultLangID;
end;

function TResFile.TLangID.GetName: string;
begin
  SetLength(Result, LOCALE_NAME_MAX_LENGTH);
  var Len := LCIDToLocaleName(FLangID, PChar(Result), LOCALE_NAME_MAX_LENGTH, LOCALE_ALLOW_NEUTRAL_NAMES);
  if Len = 0 then
    RaiseLastOSError;
  SetLength(Result, Len - 1);
end;

class operator TResFile.TLangID.Implicit(const Value: LANGID): TLangID;
begin
  Result.SetLang(0, Value);
end;

class operator TResFile.TLangID.Implicit(const Value: TLangID): LANGID;
begin
  Result := Value.FLangID;
end;

procedure TResFile.TLangID.SetLang(idx: Integer; Value: Word);
begin
  case idx of
    0: FLangID := Winapi.Windows.MAKELANGID(Value, SubLanguage);
    1: FLangID := Winapi.Windows.MAKELANGID(PrimaryLanguage, Value);
  end;
end;

procedure TResFile.TLangID.SetLangID(Primary, SubLang: Word);
begin
  FLangID := Winapi.Windows.MAKELANGID(Primary, SubLang);
end;

procedure TResFile.TLangID.SetName(Value: string);
begin
  Self.FLangID := GetLangID(Value);
end;

{ TResFile.TVersion }

function TResFile.TVersion.GetStr: string;
begin
  Result := Format('%.2d.%.2d', [Major, Minor]);
end;

function TResFile.TVersion.GetVer: Single;
begin
  Result := Major + (Min(Minor, 99) / 100)
end;

class operator TResFile.TVersion.Implicit(const Value: DWORD): TVersion;
begin
  Result.Raw := Value;
end;

class operator TResFile.TVersion.Implicit(const Value: TVersion): DWORD;
begin
  Result := Value.Raw;
end;

procedure TResFile.TVersion.SetStr(Value: string);
begin
  Version := Single.Parse(Value, TFormatSettings.Invariant);
end;

procedure TResFile.TVersion.SetVer(Value: Single);
begin
  Major := Trunc(Value);
  Minor := Trunc(Frac(Value) * 100 + 0.0001);
end;

{ TResFile.TResource }

function TResFile.TResource.DataSize: Integer;
begin
  Result := Length(Data);
end;

class operator TResFile.TResource.Initialize(out Dest: TResource);
begin
  Dest.Type_            := '';
  Dest.Name             := '';
  Dest.ResVersion.Raw   := 0;
  Dest.Flags            := $FFFF;       // TResFile.DefaultFlags;
  Dest.Language.FLangID := $FFFF;       // TResFile.DefaultLangID;
  Dest.Version.Raw      := $FFFF_FFFF;  // TResFile.DefaultVersion;
  Dest.Characteristics  := 0;
  Dest.Data             := nil;
end;

function TResFile.TResource.IsEmpty: Boolean;
begin
  Result := not Assigned(Data) and MatchStr(Name, ['', '#0']);
end;

procedure TResFile.TResource.ReadRecord(const FileData: TBytes; var Pos: Integer);
function GetPos(Size: Integer): Pointer;
  begin
    Result := PByte(FileData) + Pos;
    Inc(Pos, Size);
    if Pos > Length(FileData) then
      raise EReadError.CreateRes(@SReadPastEndOfStream) at ReturnAddress;
  end;
function GetString: string;
  begin
    Result := PWideChar(GetPos(2));
    GetPos(Length(Result) * 2);
    if Result = '' then
      raise EReadError.CreateRes(@SInvalidStringLength) at ReturnAddress;
  end;
function GetIndexString(AsType: Boolean): string;
  begin
    var DataType := pWORD(GetPos(2))^;
    if DataType = MAXWORD then begin
      var ID := pWORD(GetPos(2))^;
      if AsType and (SmallInt(ID) < Ord(High(TResFile.cDataTypes))) and (TResFile.cDataTypes[Char(ID)] <> '') then
        Result := TResFile.cDataTypes[Char(ID)]
      else
        Result := '#' + ID.ToString;
    end else begin
      Dec(Pos, 2);  // revert DataType
      Result := GetString;
    end;
  end;
procedure DataAlign;
  begin
    if Pos mod 4 <> 0 then
      Inc(Pos, 4 - Pos mod 4);
  end;
begin
  var HeaderStart   := Pos;
  var DataSize      := pDWORD(GetPos(4))^;
  var HeaderSize    := pDWORD(GetPos(4))^;

  Type_.FName       := GetIndexString(True);
  DataAlign;
  Name.FName        := GetIndexString(False);
  DataAlign;
  ResVersion.Raw    := pDWORD(GetPos(4))^;
  Flags             :=  pWORD(GetPos(2))^;
  Language.FLangID  :=  pWORD(GetPos(2))^;
  Version.Raw       := pDWORD(GetPos(4))^;
  Characteristics   := pDWORD(GetPos(4))^;
  if Pos - HeaderStart <> Integer(HeaderSize) then
    raise EReadError.CreateRes(@SReadError);

  Data := Copy(FileData, Pos, DataSize);
  GetPos(DataSize);
  DataAlign;
end;

procedure TResFile.TResource.ResolveDefaults;
begin
  if ResVersion.Raw = $FFFF_FFFF then
    ResVersion.Raw := 0;

  if Flags = $FFFF then
    Flags := TResFile.DefaultFlags;

  if Language.LangID = $FFFF then
    Language.LangID := TResFile.DefaultLangID.FLangID;
  if Language.LangID = $FFFF then
    Language.LangID := TResFile.TLangID.GetDefaultLangID.FLangID;

  if Version.Raw = $FFFF_FFFF then
    Version.Raw := TResFile.DefaultVersion.Raw;
  if Version.Raw = $FFFF_FFFF then
    Version.Raw := 0;

  //if Characteristics = $FFFF_FFFF then
  //  Characteristics = 0;
end;

procedure TResFile.TResource.WriteRecord(var FileData: TBytes; var Pos: Integer);
function GetPos(Size: Integer): Pointer;
  begin
    Result := PByte(FileData) + Pos;
    Inc(Pos, Size);
    if Pos > Length(FileData) then
      SetLength(FileData, Pos + 65_000);
  end;
procedure WriteIndexString(Value: Integer); overload;
  begin
    pWORD(GetPos(2))^ := MAXWORD;
    pWORD(GetPos(2))^ := Value;
  end;
procedure WriteIndexString(Value: string); overload;
  begin
    if StartsStr('#', Value) then begin
      pWORD(GetPos(2))^ := MAXWORD;
      pWORD(GetPos(2))^ := Value.Substring(1).ToInteger;
    end else begin
      if Value <> Value.Replace(#0, '') then
        raise EReadError.CreateRes(@SInvalidBinary) at ReturnAddress;
      var Bytes := TEncoding.Unicode.GetBytes(Value + #0);
      GetPos(0);
      Insert(Bytes, FileData, Pos);
      GetPos(Length(Bytes));
    end;
  end;
procedure DataAlign;
  begin
    if Pos mod 4 <> 0 then
      GetPos(4 - Pos mod 4);
  end;
begin
  var HeaderStart    := Pos;
  pDWORD(GetPos(4))^ := Length(Data);  // DataSize
  pDWORD(GetPos(4))^ := DWORD(-1);     // HeaderSize

  ResolveDefaults;
  if Name.isID then
    WriteIndexString(Name.GetID)  // translate named IDs
  else
    WriteIndexString(Name.FName);
  DataAlign;
  WriteIndexString(Name.FName);
  DataAlign;
  pDWORD(GetPos(4))^ := ResVersion.Raw;
   pWORD(GetPos(2))^ := Flags;
   pWORD(GetPos(2))^ := Language.LangID;
  pDWORD(GetPos(4))^ := Version.Raw;
  pDWORD(GetPos(4))^ := Characteristics;

  pDWORD(PByte(FileData) + HeaderStart + 4)^ := Pos - HeaderStart;

  GetPos(0);
  Insert(Data, FileData, Pos);
  GetPos(Length(Data));
  DataAlign;
end;

{ TResFile }

function TResFile.Add(Resource: TResource): Integer;
begin
  Resource.ResolveDefaults;
  if Resource.IsEmpty then
    Exit(-1);
  if IndexOf(Resource) >= 0 then
    raise EReadError.CreateRes(@SGenericDuplicateItem);
  Result := Length(Resources);
  Insert(Resource, Resources, Result);
end;

function TResFile.Add(&Type: TTypeOrID; Name: TNameOrID; Data: TBytes; Language: LANGID; Flags: Word): Integer;
begin
  var Resource: TResource;
  Resource.Type_            := &Type;
  Resource.Name             := Name;
  Resource.ResVersion.Raw   := 0;
  Resource.Flags            := Flags;     // $FFFF => DefaultFlags
  Resource.Language.FLangID := Language;  // $FFFF => DefaultLangID or TLangID.GetDefaultLangID
  Resource.Version.FData    := DefaultVersion.FData;
  Resource.Characteristics  := 0;
  Resource.Data             := Data;
  Result := Add(Resource);
end;

class function TResFile.CheckIsPEFile(FileName: string): Boolean;
begin
  Result := not EndsText('.res', FileName) and FileExists(FileName);
  if Result then begin
    var TestFile := TFile.OpenRead(Filename);
    try
      var MagicBytes: array[0..1] of AnsiChar;
      Result := (TestFile.Read(MagicBytes, 2) = 2) and (MagicBytes = 'PE');
    finally
      TestFile.Free;
    end;
  end;
end;

procedure TResFile.Clear;
begin
  Self.Resources := nil;
end;

class constructor TResFile.Create;
begin
  DefaultLangID.LangID := TLangID.GetDefaultLangID.LangID;
  DefaultVersion.Raw   := 0;
  DefaultFlags         := 0;
end;

class function TResFile.Create: TResFile;
begin
  Result.Clear;
end;

constructor TResFile.Create(FileName: string);
begin
  Clear;
  if FileName <> '' then
    LoadFrom(FileName);
end;

procedure TResFile.Delete(idx: Integer);
begin
  System.Delete(Resources, idx, 1);
end;

class function TResFile.EnumResLangProc(hModule: HMODULE; lpszType, lpszName: PChar; wIDLanguage: Word; lParam: NativeInt): BOOL;
var
  Type_, Name: string;
  ResStream:   TResourceStream;
  ResData:     TBytes;
begin
  if (lpszType <> nil) and (UIntPtr(lpszType) <= MAXWORD) then
    Type_ := '#' + IntPtr(lpszType).ToString else Type_ := lpszType;
  if (lpszName <> nil) and (UIntPtr(lpszName) <= MAXWORD) then
    Name  := '#' + IntPtr(lpszName).ToString else Name  := lpszName;

  if StartsStr('#', Name) then
    ResStream := TResourceStream.CreateFromID(hModule, IntPtr(lpszName), lpszType)
  else
    ResStream := TResourceStream.Create(hModule, Name, lpszType);
  try
    SetLength(ResData, ResStream.Size);
    ResStream.ReadBuffer(ResData, ResStream.Size);
    var Index := TResFile(Pointer(lParam)^).Add(Type_, Name, ResData, wIDLanguage);
    TResFile(Pointer(lParam)^).Resources[Index].SortOrder := ResStream.Memory;
  finally
    ResStream.Free;
  end;
  Result := True;
end;

class function TResFile.EnumResNameProc(hModule: HMODULE; lpszType, lpszName: PChar; lParam: NativeInt): BOOL;
begin
  if not EnumResourceLanguages(hModule, lpszType, lpszName, @EnumResLangProc, lParam) then
    RaiseLastOSError;
  Result := True;
end;

class function TResFile.EnumResTypeProc(hModule: HMODULE; lpszType: PChar; lParam: NativeInt): BOOL;
begin
  if not EnumResourceNames(HInstance, lpszType, @EnumResNameProc, lParam) then
    RaiseLastOSError;
  Result := True;
end;

function TResFile.Find(&Type: TTypeOrID): TArray<TResource>;
begin
  SetLength(Result, Length(Resources));
  var Count := 0;
  for var Index := 0 to High(Resources) do begin
    Resources[Index].ResolveDefaults;
    if Resources[Index].Type_.Name = &Type.Name then begin
      Result[Count] := Resources[Index];
      Inc(Count);
    end;
  end;
  SetLength(Result, Count);
end;

function TResFile.Find(&Type: TTypeOrID; Name: TNameOrID): TArray<TResource>;
begin
  SetLength(Result, Length(Resources));
  var Count := 0;
  for var Index := 0 to High(Resources) do begin
    Resources[Index].ResolveDefaults;
    if (Resources[Index].Type_.Name = &Type.Name) and (Resources[Index].Name.Name = Name.Name) then begin
      Result[Count] := Resources[Index];
      Inc(Count);
    end;
  end;
  SetLength(Result, Count);
end;

function TResFile.FindLang(Language: TLangID): TArray<TResource>;
begin
  if Language.LangID = $FFFF then
    Language.LangID := TResFile.DefaultLangID.FLangID;
  if Language.LangID = $FFFF then
    Language.LangID := TResFile.TLangID.GetDefaultLangID.FLangID;

  SetLength(Result, Length(Resources));
  var Count := 0;
  for var Index := 0 to High(Resources) do begin
    Resources[Index].ResolveDefaults;
    if Resources[Index].Language.LangID = Language.LangID then begin
      Result[Count] := Resources[Index];
      Inc(Count);
    end;
  end;
  SetLength(Result, Count);
end;

function TResFile.FindName(Name: TNameOrID): TArray<TResource>;
begin
  SetLength(Result, Length(Resources));
  var Count := 0;
  for var Index := 0 to High(Resources) do begin
    Resources[Index].ResolveDefaults;
    if Resources[Index].Name.Name = Name.Name then begin
      Result[Count] := Resources[Index];
      Inc(Count);
    end;
  end;
  SetLength(Result, Count);
end;

function TResFile.FindLanguages: TArray<TLangID>;
type
  TWordArray = TArray<WORD>;
var
  Found: NativeInt;
begin
  Result := nil;
  for var Index := 0 to High(Resources) do begin
    Resources[Index].ResolveDefaults;
    if not TArray.BinarySearch<{TLangID}WORD>(TWordArray(Result), Resources[Index].Language.LangID, Found) then
      Insert(Resources[Index].Language, Result, Found);
  end;
  if not Assigned(Result) then
    Result := [DefaultLangID];
end;

function TResFile.IndexOf(Resource: TResource; IgnoreLanguage: Boolean): Integer;
begin
  Result := -1;
  Resource.ResolveDefaults;
  if Resource.IsEmpty then
    Exit;
  for var idx := 0 to High(Resources) do begin
    Resources[idx].ResolveDefaults;
    if (Resources[idx].Type_.FName = Resource.Type_.FName) and (Resources[idx].Name.FName = Resource.Name.FName) then
      if Resources[idx].Language.FLangID = Resource.Language.FLangID then
        Exit(idx)
      else
        if IgnoreLanguage and (Result < 0) then
          Result := idx;
  end;
end;

function TResFile.IndexOf(&Type: TTypeOrID; Name: TNameOrID; IgnoreLanguage: Boolean; Language: LANGID): Integer;
begin
  var Resource: TResource;
  Resource.Type_    := &Type;
  Resource.Name     := Name;
  Resource.Language := Language;
  Result := IndexOf(Resource, IgnoreLanguage);
end;

class operator TResFile.Initialize(out Dest: TResFile);
begin
  Dest.Resources := nil;
end;

//function TResFile.List(Language: TLangID): TArray<TResource>;
//function TResFile.List(Name: TNameOrID): TArray<TResource>;
//function TResFile.List(&Type: TTypeOrID): TArray<TResource>;
//function TResFile.List(&Type: TTypeOrID; Name: TNameOrID): TArray<TResource>;
//function TResFile.ListLanguages: TArray<TLangID>;

procedure TResFile.LoadFrom(FileName: string);
begin
  if CheckIsPEFile(FileName) then begin
    var Module := LoadLibraryEx(PChar(FileName), 0, LOAD_LIBRARY_AS_DATAFILE);
    if Module = 0 then
      RaiseLastOSError;
    try
      Clear;
      if not EnumResourceTypes(Module, @EnumResTypeProc, IntPtr(@Self)) then
        RaiseLastOSError;
      SortEnumResources;
    finally
      FreeLibrary(Module);
    end;
  end else
    LoadFrom(TFile.ReadAllBytes(FileName));
end;

procedure TResFile.LoadFrom(FileData: TBytes);
begin
  FileData := Copy(FileData);
  var idx  := 0;
  var Pos  := 0;
  try
    SetLength(Resources, 10);
    Resources[0].ReadRecord(FileData, Pos);  // actually a RESOURCEHEADER struct, but it also simply corresponds to an empty resource entry
    if (Pos <> 32{4+4 + 4+4 + 4+2+2+4+4}) or not Resources[0].IsEmpty then
      raise EReadError.CreateRes(@SInvalidImage);  // 16 bit resource files are not supported
    while Pos + 2 * SizeOf(DWORD) < Length(FileData) do begin
      if idx >= High(Resources) then
        SetLength(Resources, Length(Resources) + 25);
      Resources[idx].ReadRecord(FileData, Pos);
      if not Resources[idx].IsEmpty then
        Inc(idx);
    end;
  finally
    SetLength(Resources, idx);
  end;
end;

procedure TResFile.LoadMine;
begin
  Clear;
  if not EnumResourceTypes(HInstance, @EnumResTypeProc, IntPtr(@Self)) then
    RaiseLastOSError;
  SortEnumResources;
end;

function TResFile.ResCount: Integer;
begin
  Result := Length(Resources);
end;

class function TResFile.ResNameToStr(ResName: PChar): string;
begin
  if (ResName <> nil) and (UIntPtr(ResName) <= MAXWORD) then
    Result := '#' + IntPtr(ResName).ToString
  else
    Result := ResName;
end;

class function TResFile.ResTypeToStr(ResType: PChar): string;
begin
  if (ResType <> nil) and (UIntPtr(ResType) <= MAXWORD) then
    if (IntPtr(ResType) <= Ord(High(cDataTypes))) and (cDataTypes[Char(ResType)] <> '') then
      Result := cDataTypes[Char(ResType)]
    else
      Result := '#' + IntPtr(ResType).ToString
  else
    Result := ResType;
end;

function TResFile.SaveTo: TBytes;
begin
  var Pos := 0;
  Result  := nil;
  try
    var Header := Default(TResource);  // empty record as marker for 32 bit resources
    Header.Flags            := 0;
    Header.Language.FLangID := 0;
    Header.Version.Raw      := 0;
    Header.WriteRecord(Result, Pos);   // actually a RESOURCEHEADER struct, but it also simply corresponds to an empty resource entry
    for var idx := 0 to High(Resources) do begin
      if IndexOf(Resources[idx]) < idx then
        raise EReadError.CreateRes(@SGenericDuplicateItem);
      if not Resources[idx].IsEmpty then
        Resources[idx].WriteRecord(Result, Pos);
    end;
  finally
    SetLength(Result, Pos);
  end;
end;

procedure TResFile.SaveTo(FileName: string; RemoveMissingFromPEFiles: Boolean);
begin
  if CheckIsPEFile(FileName) then begin
    var Success := True;
    var Module  := BeginUpdateResource(PChar(Filename), RemoveMissingFromPEFiles);
    if Module = 0 then
      RaiseLastOSError;
    try
      for var Index := 0 to High(Resources) do begin
        {$REGION 'resource info'}
        Resources[Index].ResolveDefaults;
        var Resource := Resources[Index];
        var ResType  := Resource.Type_.GetRI;
        var ResName  := Resource.Name.GetNI;
        var ResLang  := Resource.Language.FLangID;
        var ResData  := Pointer(Resource.Data);
        var ResSize  := Resource.DataSize;
        {$ENDREGION}
        {$REGION 'check update'}
        var DoUpdate := RemoveMissingFromPEFiles;  // all resources were deleted during BeginUpdate
        if not RemoveMissingFromPEFiles then begin
          var ResInfo := FindResourceEx(Module, ResName, ResType, ResLang);
          if ResInfo <> 0 then begin
            var HGlobal := LoadResource(Module, ResInfo);
            if HGlobal <> 0 then
              try
                var _ResSize := SizeOfResource(Module, ResInfo);
                var _ResData := LockResource(HGlobal);
                if (_ResData = nil) and (_ResSize <> 0) then
                  DoUpdate := True  // or RaiseLastOSError
                else
                  DoUpdate := (ResSize <> Integer(_ResSize)) or CompareMem(ResData, _ResData, ResSize);
              finally
                UnlockResource(HGlobal);
                FreeResource(HGlobal);
              end
            else
              DoUpdate := True;  // or RaiseLastOSError
          end else
            DoUpdate := True;
        end;
        {$ENDREGION}
        if DoUpdate then
          if not UpdateResource(Module, ResType, ResName, ResLang, ResData, ResSize) then begin
            Success := False;
            Break;  // or RaiseLastOSError
          end;
      end;
    finally
      if not EndUpdateResource(Module, not Success) then
        RaiseLastOSError;
    end;
  end else
    TFile.WriteAllBytes(FileName, SaveTo);
end;

procedure TResFile.SortEnumResources;
begin
  // Was loaded sorted by type, but I want them in the order they were saved/created.
  // EnumResourceTypes > EnumResourceNames > EnumResourceLanguages > FindResource > LoadResource
  for var a := 0 to High(Resources) - 1 do
    for var b := a + 1 to High(Resources) do
      if IntPtr(Resources[a].SortOrder) > IntPtr(Resources[b].SortOrder) then begin
        var Temp     := Resources[a];
        Resources[a] := Resources[b];
        Resources[b] := Temp;
      end;
end;

end.

