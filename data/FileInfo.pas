unit FileInfo;

interface

uses Windows, ActiveX, Classes;

const
  FMTID_SummaryInformation: TGUID =     '{F29F85E0-4FF9-1068-AB91-08002B27B3D9}';
  FMTID_DocSummaryInformation: TGUID = '{D5CDD502-2E9C-101B-9397-08002B2CF9AE}';
  FMTID_UserDefinedProperties: TGUID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
  FMTID_ImageSummaryInformation: TGUID = '{6444048F-4C8B-11D1-8B70-080036B11A03}';
  FMTID_InternetSite: TGUID = '{000214A1-0000-0000-C000-000000000046}';
  FMTID_Music: TGUID = '{56A3372E-CE9C-11D2-9F0E-006097C686F6}';
  FMTID_Audio: TGUID = '{64440490-4C8B-11D1-8B70-080036B11A03}';
  FMTID_Video: TGUID = '{64440491-4C8B-11D1-8B70-080036B11A03}';
  FMTID_MediaFile: TGUID = '{64440492-4C8B-11D1-8B70-080036B11A03}';

  CLSID_PropertiesUI: TGUID = '{d912f8cf-0396-4915-884e-fb425d32943b}';
  SID_IPropertyUI = '{757a7d9f-919a-4118-99d7-dbb208c8cc66}';
  IID_IPropertyUI: TGUID = SID_IPropertyUI;
  IID_IPropertySetStorage: TGUID = '{0000013A-0000-0000-C000-000000000046}';

  PropTypes: array[0..8] of TGUID = (
    (*FMTID_SummaryInformation*)'{F29F85E0-4FF9-1068-AB91-08002B27B3D9}',
    (*FMTID_DocSummaryInformation*)'{D5CDD502-2E9C-101B-9397-08002B2CF9AE}',
    (*FMTID_UserDefinedProperties*)'{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
    (*FMTID_ImageProperties*)'{14B81DA1-0135-4d31-96d9-6CBFC9671A99}',
    (*FMTID_InternetSite*)'{000214A1-0000-0000-C000-000000000046}',
    (*FMTID_Music*)'{56A3372E-CE9C-11D2-9F0E-006097C686F6}',
    (*FMTID_Video*)'{64440491-4C8B-11D1-8B70-080036B11A03}',
    (*FMTID_Audio*)'{64440490-4C8B-11D1-8B70-080036B11A03}',
    (*FMTID_MediaFile*)'{64440492-4C8B-11D1-8B70-080036B11A03}');

  //Indicates that the file must not be a compound file.
  //This element is only valid when using the StgCreateStorageEx
  //or StgOpenStorageEx functions to access the NTFS file system
  //implementation of the IPropertySetStorage interface.
  //Therefore, these functions return an error if the riid
  //parameter does not specify the IPropertySetStorage interface,
  //or if the specified file is not located on an NTFS file system volume.
  STGFMT_FILE = 3;

  //Indicates that the system will determine the file type and
  //use the appropriate structured storage or property set implementation.
  //This value cannot be used with the StgCreateStorageEx function.
  STGFMT_ANY = 4;

  // Summary Information
  PID_TITLE        = 2;
  PID_SUBJECT      = 3;
  PID_AUTHOR       = 4;
  PID_KEYWORDS     = 5;
  PID_COMMENTS     = 6;
  PID_TEMPLATE     = 7;
  PID_LASTAUTHOR   = 8;
  PID_REVNUMBER    = 9;
  PID_EDITTIME     = 10;
  PID_LASTPRINTED  = 11;
  PID_CREATE_DTM   = 12;
  PID_LASTSAVE_DTM = 13;
  PID_PAGECOUNT    = 14;
  PID_WORDCOUNT    = 15;
  PID_CHARCOUNT    = 16;
  PID_THUMBNAIL    = 17;
  PID_APPNAME      = 18;
  PID_SECURITY     = 19;

  // Document Summary Information
  PID_CATEGORY     = 2;
  PID_PRESFORMAT   = 3;
  PID_BYTECOUNT    = 4;
  PID_LINECOUNT    = 5;
  PID_PARCOUNT     = 6;
  PID_SLIDECOUNT   = 7;
  PID_NOTECOUNT    = 8;
  PID_HIDDENCOUNT  = 9;
  PID_MMCLIPCOUNT  = 10;
  PID_SCALE        = 11;
  PID_HEADINGPAIR  = 12;
  PID_DOCPARTS     = 13;
  PID_MANAGER      = 14;
  PID_COMPANY      = 15;
  PID_LINKSDIRTY   = 16;
  PID_CHARCOUNT2   = 17;

  // FMTID_ImageSummaryInformation
  PIDISI_FILETYPE  = 2;  // VT_LPWSTR
  PIDISI_CX        = 3;  // VT_UI4
  PIDISI_CY        = 4;  // VT_UI4
  PIDISI_RESOLUTIONX = 5;  // VT_UI4
  PIDISI_RESOLUTIONY = 6;  // VT_UI4
  PIDISI_BITDEPTH = 7;  // VT_UI4
  PIDISI_COLORSPACE = 8;  // VT_LPWSTR
  PIDISI_COMPRESSION = 9;  // VT_LPWSTR
  PIDISI_TRANSPARENCY = 10;  // VT_UI4
  PIDISI_GAMMAVALUE = 11;  // VT_UI4
  PIDISI_FRAMECOUNT = 12;  // VT_UI4
  PIDISI_DIMENSIONS = 13;  // VT_LPWSTR

  // FMTID_Music
  PIDSI_ARTIST    = 2;
  PIDSI_SONGTITLE = 3;
  PIDSI_ALBUM     = 4;
  PIDSI_YEAR      = 5;
  PIDSI_COMMENT   = 6;
  PIDSI_TRACK     = 7;
  PIDSI_GENRE     = 11;
  PIDSI_LYRICS    = 12;

  // FMTID_Video
  PIDVSI_STREAM_NAME   = $00000002; // "StreamName", VT_LPWSTR
  PIDVSI_FRAME_WIDTH   = $00000003; // "FrameWidth", VT_UI4
  PIDVSI_FRAME_HEIGHT  = $00000004; // "FrameHeight", VT_UI4
  PIDVSI_TIMELENGTH    = $00000007; // "TimeLength", VT_UI4, milliseconds
  PIDVSI_FRAME_COUNT   = $00000005; // "FrameCount". VT_UI4
  PIDVSI_FRAME_RATE    = $00000006; // "FrameRate", VT_UI4, frames/millisecond
  PIDVSI_DATA_RATE     = $00000008; // "DataRate", VT_UI4, bytes/second
  PIDVSI_SAMPLE_SIZE   = $00000009; // "SampleSize", VT_UI4
  PIDVSI_COMPRESSION   = $0000000A; // "Compression", VT_LPWSTR
  PIDVSI_STREAM_NUMBER = $0000000B; // "StreamNumber", VT_UI2

  // FMTID_AudioSummaryInformation property identifiers
  PIDASI_FORMAT        = $00000002; // VT_BSTR
  PIDASI_TIMELENGTH    = $00000003; // VT_UI4, milliseconds
  PIDASI_AVG_DATA_RATE = $00000004; // VT_UI4,  Hz
  PIDASI_SAMPLE_RATE   = $00000005; // VT_UI4,  bits
  PIDASI_SAMPLE_SIZE   = $00000006; // VT_UI4,  bits
  PIDASI_CHANNEL_COUNT = $00000007; // VT_UI4
  PIDASI_STREAM_NUMBER = $00000008; // VT_UI2
  PIDASI_STREAM_NAME   = $00000009; // VT_LPWSTR
  PIDASI_COMPRESSION   = $0000000A; // VT_LPWSTR

function StgOpenStorageEx(
  const pwcsName: POleStr;  {Pointer to the path of the file containing storage object}
  grfMode: LongInt;         {Specifies the access mode for the object}
  stgfmt: DWORD;            {Specifies the storage file format}
  grfAttrs: DWORD;          {Reserved; must be zero}
  pStgOptions: Pointer;     {Address of STGOPTIONS pointer}
  reserved2: Pointer;       {Reserved; must be zero}
  riid: PGUID;              {Specifies the GUID of the interface pointer}
  out stgOpen: IStorage     {Address of an interface pointer}
  ): HResult; stdcall; external 'ole32.dll';

function ListFileInfo(const AFileName: WideString; const AProps: TStrings): Boolean;
function SummaryIDToString(const ePID: Cardinal): string; inline;
function DocumentIDToString(const ePID: Cardinal): string; inline;
function UserIDToString(const ePID: Cardinal): string; inline;
function ImageInfoIDToString(const ePID: Cardinal): string; inline;
function MusicInfoIDToString(const ePID: Cardinal): string; inline;
function VideoInfoIDToString(const ePID: Cardinal): string; inline;
function AudioInfoIDToString(const ePID: Cardinal): string; inline;
function GetPropertyDisplayName(const AFMTID: TGUID; const APID: Cardinal): string;

type
  IPropertyUI = interface(IUnknown)
    [SID_IPropertyUI]
    function ParsePropertyName(pszName : LPCWSTR; var pfmtid : FMTID; var ppid : PROPID; var chEaten : ulong) : hresult; stdcall;
    function GetCannonicalName(fmtid : TGUID; pid : PROPID; pwszText : LPWSTR; cchText : DWORD) : hresult; stdcall;
    function GetDisplayName(fmtid : TGUID; pid : PROPID; flags : integer; pwszText : LPWSTR; cchText : dword) : hresult; stdcall;
    function GetPropertyDescription(fmtid : TGUID; pid : PROPID; pwszText : LPWSTR; cchText : DWORD) : hresult; stdcall;
    function GetDefaultWidth(fmtid : TGUID; pid : PROPID; var pcxChars : ULONG) : hresult; stdcall;
    function GetFlags(fmtid : TGUID; pid : PROPID; var pFlags : integer) : hresult; stdcall;
    function FormatForDisplay(fmtid : TGUID; pid : PROPID; const pvar : PROPVARIANT; flags : integer; pwszText : LPWSTR; cchText : dword) : hresult; stdcall;
    function GetHelpInfo(fmtid : TGUID; pid : PROPID; pwszHelpFile : LPWSTR; cch : DWORD; var puHelpID : uint) : hresult; stdcall;
  end;

  TIDToString = function(const AID: Cardinal): string;
const
  IDToString: array[0..8] of TIDToString = (
    SummaryIDToString, DocumentIDToString, UserIDToString, ImageInfoIDToString,
    UserIDToString, MusicInfoIDToString, VideoInfoIDToString,
    AudioInfoIDToString, UserIDToString);

implementation

uses SysUtils, ComObj;

function SummaryIDToString(const ePID: Cardinal): string;
begin
  case ePID of
    PID_TITLE: Result := 'Title';
    PID_SUBJECT: Result := 'Subject';
    PID_AUTHOR: Result := 'Author';
    PID_KEYWORDS: Result := 'Keywords';
    PID_COMMENTS: Result := 'Comments';
    PID_TEMPLATE: Result := 'Template';
    PID_LASTAUTHOR: Result := 'Last Saved By';
    PID_REVNUMBER: Result := 'Revision Number';
    PID_EDITTIME: Result := 'Total Editing Time';
    PID_LASTPRINTED: Result := 'Last Printed';
    PID_CREATE_DTM: Result := 'Create Time/Date';
    PID_LASTSAVE_DTM: Result := 'Last Saved Time/Date';
    PID_PAGECOUNT: Result := 'Number of Pages';
    PID_WORDCOUNT: Result := 'Number of Words';
    PID_CHARCOUNT: Result := 'Number of Characters';
    PID_THUMBNAIL: Result := 'Thumbnail';
    PID_APPNAME: Result := 'Creating Application';
    PID_SECURITY: Result := 'Security';
  else
    Result := '$' + IntToHex(ePID, 8);
  end;
end;

function DocumentIDToString(const ePID: Cardinal): string;
begin
  case ePID of
    PID_CATEGORY: Result := 'Category';
    PID_PRESFORMAT: Result := 'Format';
    PID_BYTECOUNT: Result := 'Byte Count';
    PID_LINECOUNT: Result := 'Line Count';
    PID_PARCOUNT: Result := 'Paragraph Count';
    PID_SLIDECOUNT: Result := 'Slide Count';
    PID_NOTECOUNT: Result := 'Note Count';
    PID_HIDDENCOUNT: Result := 'Hidden Slides';
    PID_MMCLIPCOUNT: Result := 'Clip Count';
    PID_SCALE: Result := 'Scaled';
    PID_HEADINGPAIR: Result := 'Heading Pair';
    PID_DOCPARTS: Result := 'Document Parts';
    PID_MANAGER: Result := 'Manager';
    PID_COMPANY: Result := 'Company';
    PID_LINKSDIRTY: Result := 'Link Dirty';
    PID_CHARCOUNT2: Result := 'Character Count';
  else
    Result := '$' + IntToHex(ePID, 8);
  end;
end;

function UserIDToString(const ePID: Cardinal): string;
begin
  Result := '$' + IntToHex(ePID, 8);
end;

function ImageInfoIDToString(const ePID: Cardinal): string;
begin
  case ePID of
    PIDISI_FILETYPE: Result := 'File Type';
    PIDISI_CX: Result := 'Width';
    PIDISI_CY: Result := 'Height';
    PIDISI_RESOLUTIONX: Result := 'Horizontal Resolution';
    PIDISI_RESOLUTIONY: Result := 'Vertical Resolution';
    PIDISI_BITDEPTH: Result := 'Bit Depth';
    PIDISI_COLORSPACE: Result := 'ColorSpace';
    PIDISI_COMPRESSION: Result := 'Compression';
    PIDISI_TRANSPARENCY: Result := 'Transparency';
    PIDISI_GAMMAVALUE: Result := 'Gamma';
    PIDISI_FRAMECOUNT: Result := 'Frame Count';
    PIDISI_DIMENSIONS: Result := 'Dimensions';
  else
    Result := '$' + IntToHex(ePID, 8);
  end;
end;

function MusicInfoIDToString(const ePID: Cardinal): string;
begin
  case ePID of
    PIDSI_ARTIST: Result := 'Artist';
    PIDSI_SONGTITLE: Result := 'Title';
    PIDSI_ALBUM: Result := 'Album';
    PIDSI_YEAR: Result := 'Year';
    PIDSI_COMMENT: Result := 'Comment';
    PIDSI_TRACK: Result := 'Track';
    PIDSI_GENRE: Result := 'Genre';
    PIDSI_LYRICS: Result := 'Lyrics';
  else
    Result := '$' + IntToHex(ePID, 8);
  end;
end;

function VideoInfoIDToString(const ePID: Cardinal): string;
begin
  case ePID of
    PIDVSI_STREAM_NAME: Result := 'StreamName';
    PIDVSI_FRAME_WIDTH: Result := 'FrameWidth';
    PIDVSI_FRAME_HEIGHT: Result := 'FrameHeight';
    PIDVSI_TIMELENGTH: Result := 'TimeLength';
    PIDVSI_FRAME_COUNT: Result := 'FrameCount';
    PIDVSI_FRAME_RATE: Result := 'FrameRate';
    PIDVSI_DATA_RATE: Result := 'DataRate';
    PIDVSI_SAMPLE_SIZE: Result := 'SampleSize';
    PIDVSI_COMPRESSION: Result := 'Compression';
    PIDVSI_STREAM_NUMBER: Result := 'StreamNumber';
  else
    Result := '$' + IntToHex(ePID, 8);
  end;
end;

function AudioInfoIDToString(const ePID: Cardinal): string;
begin
  case ePID of
    PIDASI_FORMAT: Result := 'Format';
    PIDASI_TIMELENGTH: Result := 'Length';
    PIDASI_AVG_DATA_RATE: Result := 'Data Rate';
    PIDASI_SAMPLE_RATE: Result := 'Sample Rate';
    PIDASI_SAMPLE_SIZE: Result := 'Sample Size';
    PIDASI_CHANNEL_COUNT: Result := 'Channels';
    PIDASI_STREAM_NUMBER: Result := 'Streams';
    PIDASI_STREAM_NAME: Result := 'Stream Name';
    PIDASI_COMPRESSION: Result := 'Compression';
  else
    Result := '$' + IntToHex(ePID, 8);
  end;
end;

function PropVariantToString(const Prop: TPropVariant): string;
var
  TheDate: TFileTime;
  SysDate: TSystemTime;
begin
  case Prop.vt of
    VT_EMPTY, VT_NULL, VT_VOID: Result := '';
    VT_I2: Result := IntToStr(Prop.iVal);
    VT_I4, VT_INT: Result := IntToStr(Prop.lVal);
    VT_R4: Result := FloatToStr(Prop.fltVal);
    VT_R8: Result := FloatToStr(Prop.dblVal);
    VT_BSTR: Result := Prop.bstrVal;
    VT_ERROR: Result := 'Error code: ' + IntToStr(Prop.sCode);
    VT_BOOL:
      if Prop.boolVal = False then
        Result := 'No'
      else
        Result := 'Yes';
    VT_UI1: Result := IntToStr(Prop.bVal);
    VT_UI2: Result := IntToStr(Prop.uiVal);
    VT_UI4, VT_UINT: Result := IntToStr(Prop.ulVal);
    VT_I8: Result := IntToStr(Int64(Prop.hVal));
    VT_UI8: Result := IntToStr(Int64(Prop.uhVal));
    VT_LPSTR: Result := Prop.pszVal;
    VT_LPWSTR: Result := Prop.pwszVal;
    VT_FILETIME:
    begin
      FileTimeToLocalFileTime(TFileTime(Prop.date), TheDate);
      FileTimeToSystemTime(TheDate, SysDate);
      if SysDate.wYear < 1980 then
        Result := TimeToStr(SystemTimeToDateTime(SysDate))
      else
        Result := DateTimeToStr(SystemTimeToDateTime(SysDate));
    end;
    VT_CLSID: if Prop.puuid <> nil then Result := GuidToString(Prop.puuid^);
  else
    Result := '';
  end;
end;

function GetPropertyDisplayName(const AFMTID: TGUID; const APID: Cardinal): string;
var
  LPUI: IPropertyUI;
  LWName: PWideChar;
  LCName: Cardinal;
begin
  Result := '';
  CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
  try
    if CoCreateInstance(CLSID_PropertiesUI, nil, CLSCTX_INPROC, IID_IPropertyUI, LPUI) <> S_OK then exit;
    LCName := 4096;
    LWName := CoTaskMemAlloc(LCName);
    try
      if LPUI.GetPropertyDescription(AFMTID, APID, LWName, (LCName-2) div 2) = S_OK then
        Result := LWName;
    finally
      CoTaskMemFree(LWName);
    end;
  finally
    CoUninitialize;
  end;
end;

function ListFileInfo(const AFileName: WideString; const AProps: TStrings): Boolean;
var
  I: Integer;
  PropSetStg: IPropertySetStorage;
  PropSpec: array of TPropSpec;
  PropSpecNames: array of string;
  PropStg: IPropertyStorage;
  PropVariant: array of TPropVariant;
  LRes: HResult;
  S: string;
  PropName: string;
  PPropName: POleStr;
  Storage: IStorage;
  PropEnum: IEnumSTATPROPSTG;
  hr: HResult;
  PropStat: STATPROPSTG;
  K: integer;
  PropTypeIdx: Integer;
  PropID: TPropID;
  FormatID: TGUID;
begin
  Result := false;
  CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
  try
    hr := StgOpenStorageEx(PWideChar(AFileName), STGM_READ or STGM_SHARE_DENY_WRITE,
      STGFMT_ANY, 0, nil, nil, @IID_IPropertySetStorage, Storage);
    if hr <> S_OK then
      hr := StgOpenStorage(PWideChar(AFileName), nil, STGM_READ or STGM_SHARE_DENY_WRITE, nil, 0, Storage);
    if hr <> S_OK then exit;
    PropSetStg := Storage as IPropertySetStorage;
    for PropTypeIdx := 0 to High(PropTypes) do
    begin
      FormatID := PropTypes[PropTypeIdx];
      hr := PropSetStg.Open(FormatID,
        STGM_READ or STGM_SHARE_EXCLUSIVE, PropStg);
      if hr <> S_OK then Continue;
      OleCheck(PropStg.Enum(PropEnum));
      Result := true;
      I := 0;
      hr := PropEnum.Next(1, PropStat, nil);
      while hr = S_OK do
      begin
        Inc(I);
        SetLength(PropSpec, I);
        SetLength(PropSpecNames, I);
        PropSpec[I-1].ulKind := PRSPEC_PROPID;
        PropSpec[I-1].propid := PropStat.propid;
        PropName := PropStat.lpwstrName;
        if PropName = '' then
        begin
          PropID := PropSpec[I-1].propid;
          hr := PropStg.ReadPropertyNames(1, @PropID, @PPropName);
          if hr = S_OK then
          begin
            PropName := PPropName;
            CoTaskMemFree(PPropName);
          end;
        end;
        PropSpecNames[I-1] := PropName;
        hr := PropEnum.Next(1, PropStat, nil);
      end;
      SetLength(PropVariant, I);
      LRes := PropStg.ReadMultiple(I, @PropSpec[0], @PropVariant[0]);
      if LRes <> S_OK then Exit;
      for K := 0 to I - 1 do
      begin
        S := PropVariantToString(PropVariant[K]);
        if S <> '' then
        begin
          PropName := PropSpecNames[K];
          (*if PropName = '' then
            PropName := GetPropertyDisplayName(FormatID, PropSpec[K].propid);*)
          if PropName = '' then
            PropName := IDToString[PropTypeIdx](PropSpec[K].propid);
          AProps.Add(PropName + '=' + S);
        end;
      end;
    end;
  finally
    CoUninitialize;
  end;
end;

end.
