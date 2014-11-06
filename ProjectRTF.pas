/// text and RTF writing classes
// - this unit is part of SynProject, under GPL 3.0 license; version 1.17
unit ProjectRTF;

(*
    This file is part of SynProject.

    Synopse SynProject. Copyright (C) 2008-2011 Arnaud Bouchez
      Synopse Informatique - http://synopse.info

    SynProject is free software; you can redistribute it and/or modify it
    under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 3 of the License, or (at
    your option) any later version.

    SynProject is distributed in the hope that it will be useful, but
    WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU Lesser General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with SynProject. If not, see <http://www.gnu.org/licenses/>.

  Revision 1.13
  - TRTF is now a TClass, with virtual methods to allow not only RTF backend

  Revision 1.18
  - new THTML backend, to write HTML content

*)

interface

{.$define DIRECTPREVIEW}
{ if defined, documents can be previewed then exported as PDF directly }

{$define DIRECTEXPORTTOWORD}
{ if defined, documents are directly converted to .doc }

uses
  {$ifdef DIRECTPREVIEW}mORMotReport,{$endif}
  SysUtils, ProjectCommons, Classes;

{
   TStringWriter: fast buffered output

  var s: string;
      c: TStringWriter; // automatic Garbage Collector
  begin
    s := c.Init.Add('<').Add(HexChars,4).Add('>').
      AddXML('Essai & Deux').Add('</').AddCardinal(100).Add('>').Data;
    //  s='<0123>Essai &amp; Deux</100>'

  TRTF: specialized RTF writer
}

type
  PStringWriter = ^TStringWriter;
  TStringWriter = object
  public
    len: integer;
  private
    // tmp variable used in Add()
    sLen: integer;
    fGrowby: integer;
    fData: string;
    procedure SetData(const Value: string);
    function GetData: string;
    // returns fData without SetLength(fData,len)
    function GetDataPointer: pointer;
  public
    function Init: PStringWriter; overload;
    function Init(GrowBySize: integer): PStringWriter; overload;
    function Init(const Args: array of const): PStringWriter; overload;
    procedure SaveToStream(aStream: TStream); // contains reset (len := 0)
    procedure SaveToFile(const aFileName: String);
    procedure SaveToWriter(const aWriter: TStringWriter);
    function GetPortion(aPos, aLen: integer): string;
    procedure MovePortion(sourcePos, destPos, aLen: integer);

    procedure Write(const buf; bufLen: integer); overload;
    procedure Write(Value: integer); overload;
    procedure Write(const s: string); overload;

    function Add(c: char): PStringWriter; overload;
    function Add(p: pChar; pLen: integer): PStringWriter; overload;
    function Add(const s: string): PStringWriter; overload;
    function Add(const s1,s2: string): PStringWriter; overload;
    function Add(const s: array of string): PStringWriter; overload;
    function Add(const Format: string; const Args: array of const): PStringWriter; overload;
    function AddWithoutPeriod(const s: string; FirstCharLower: boolean = false): PStringWriter; overload;
    function AddStringOfChar(ch: char; Count: integer): PStringWriter; overload;
    function RtfBackSlash(Text: string;
      trimLastBackSlash: boolean = false): PStringWriter; // '\' -> '\\'

    function AddArray(const Args: array of const): PStringWriter;
    function AddCopy(const s: string; Index, Count: Integer): PStringWriter;
    function AddPChar(p: pChar): PStringWriter; overload;
    function AddByte(value: byte): PStringWriter;
    function AddWord(value: word): PStringWriter;
    function AddInteger(const value: integer): PStringWriter;
    // use CardinalToStrBuf()
    function AddCardinal(value: cardinal): PStringWriter;
    // AddCardinal6('+',1234)=Add('+ 001234 ')
    function AddCardinal6(cmd: char; value: cardinal; withLineNumber: boolean): PStringWriter;
    function AddInt64(value: Int64): PStringWriter;
    function AddHex32(value: cardinal): PStringWriter;
    function AddHex64(value: Int64): PStringWriter;
    function AddHexBuffer(value: PByte; Count: integer): PStringWriter;
    function AddShort(const s: shortstring): PStringWriter; overload;
    function AddShort(const s1,s2: shortstring): PStringWriter; overload;
    function AddShort(const s: array of shortstring): PStringWriter; overload;
    // no Grow test: faster
    procedure AddShortNoGrow(const s: shortstring);
    function AddCRLF: PStringWriter;

    function IsLast(const s: string): boolean; // if last data was s -> true
    function DeleteLast(const s: string): boolean; // if last data was s -> delete
    function EnsureLast(const s: string): PStringWriter; // if last data is not s -> add(s)
    function RtfValid: boolean; // return true if count of { = count of }

    property Data: string read GetData write SetData;
    property DataPointer: pointer read GetDataPointer;
  end;

  TLastRTF = (lastRtfNone, lastRtfText, lastRtfTitle, lastRtfCode,
      lastRtfImage, lastRtfCols, lastRtfList);
  TTitleLevel = array[0..6] of integer;
  TProjectWriterClass = class of TProjectWriter;
  TSaveFormat = (fNoSave,fDoc,fPdf,fHtml);

  TProjectLayout = record
    Page: record
      Width, Height: integer;
    end;
    Margin: record
      Left, Right, Top, Bottom: integer;
    end;
  end;

  TProjectWriter = class
  protected
    FontSize: integer; // font size
    WR: TStringWriter;
    fLast: TLastRTF;
    fLastWasRtfPage: boolean;
    fLastWasRtfPageInRtfTitle: boolean;
    fBookmarkInRtfTitle: string;
    fInRtfTitle: string;
    fLandscape, fKeyWordsComment: boolean;
    fStringPlain: boolean;
    // set by RtfCols():
    fCols: string; // string to be added before any row
    fColsCount: integer;
    procedure RtfKeywords(line: string; const KeyWords: array of string; aFontSize: integer = 80); virtual; abstract;
    procedure SetLast(const Value: TLastRTF); virtual; abstract;
    procedure SetLandscape(const Value: boolean); virtual;
    constructor InternalCreate; virtual;
  public
    Layout: TProjectLayout;
    Width: integer; // paragraph width in twips
    TitleWidth: integer;  // title indetation gap (default 0 twips)
    IndentWidth: integer; // list or title first line indent (default 240 twips)
    TitleLevel: TTitleLevel;
    TitleLevelCurrent: integer;
    LastTitleBookmark: string;
    MaxTitleOutlineLevel: integer;
    PicturePath: string;
    DestPath: string;
    TitlesList: TStringList; // <>nil -> RtfTitle() add one (TitleList.Free in Caller)
    TitleFlat: boolean; // true -> Titles are all numerical with no big sections
    FullTitleInTableOfContent: boolean; // true -> Title contains also \line...
    ListLine: boolean; // true -> \line, not \par in RtfList()
    FileName: TFileName;
    HandlePages: boolean;
    constructor Create(const aLayout: TProjectLayout;
      aDefFontSizeInPts: integer; // in points
      aCodePage: integer = 1252;
      aDefLang: integer = $0409; // french is $040c (1036)
      aLandscape: boolean = false; aCloseManualy: boolean = false;
      aTitleFlat: boolean = false; const aF0: string = 'Calibri';
      const aF1: string = 'Consolas'; aMaxTitleOutlineLevel: integer=5); virtual;
    function Clone: TProjectWriter; virtual;
    // if CloseManualy was true
    procedure InitClose; virtual; abstract;
    procedure SaveToFile(Format: TSaveFormat; OldWordOpen: boolean); virtual; abstract;
    function AddRtfContent(const s: string): TProjectWriter;  overload; virtual; abstract;
    function AddRtfContent(const fmt: string; const params: array of const): TProjectWriter; overload;
    function AddWithoutPeriod(const s: string; FirstCharLower: boolean = false): TProjectWriter; 
    // set Last := lastText -> update any pending {}
    function RtfText: TProjectWriter;
    procedure RtfList(line: string); virtual; abstract;
    function RtfLine: TProjectWriter; virtual; abstract;
    procedure RtfPage; virtual; abstract;
    function RtfPar: TProjectWriter; virtual; abstract;
    function RtfImageString(const Image: string; // 'SAD-4.1-functions.png 1101x738 85%'
      const Caption: string; WriteBinary: boolean; perc: integer): string; virtual; abstract;
    function RtfImage(const Image: string; const Caption: string = '';
      WriteBinary: boolean=true; const RtfHead: string='\li0\fi0\qc'): TProjectWriter; virtual; abstract;
    /// if line is displayed as code, add it and return true
    function RtfCode(const line: string): boolean;
    procedure RtfPascal(const line: string; aFontSize: integer = 80);
    procedure RtfDfm(const line: string; aFontSize: integer = 80);
    procedure RtfC(const line: string);
    procedure RtfCSharp(const line: string);
    procedure RtfListing(const line: string);
    procedure RtfSgml(const line: string);
    procedure RtfModula2(const line: string);
    procedure RtfColLine(const line: string); virtual;
    procedure RtfCols(const ColXPos: array of integer; FullWidth: integer;
      VertCentered, withBorder: boolean; const RowFormat: string = ''); virtual; abstract;
    procedure RtfColsHeader(const text: string); virtual; abstract;
    // RowFormat can be '\trkeep' e.g.
    procedure RtfColsPercent(ColWidth: array of integer;
      VertCentered, withBorder: boolean;
      NormalIndent: boolean = false;
      RowFormat: string = ''); virtual;
    // left=top, middle=center, right=bottom
    // must have been created with VertCentered=true
    procedure RtfColVertAlign(ColIndex: Integer; Align: TAlignment; DrawBottomLine: boolean=False);
    procedure RtfRow(const Text: array of string; lastRow: boolean=false); // after RtfCols
      virtual; abstract;
    function RtfColsEnd: TProjectWriter; virtual; abstract;
    procedure RtfEndSection; virtual; abstract;
    function RtfParDefault: TProjectWriter; virtual; abstract;
    procedure RtfHeaderBegin(aFontSize: integer); virtual; abstract;
    procedure RtfHeaderEnd; virtual; abstract;
    procedure RtfFooterBegin(aFontSize: integer); virtual; abstract;
    procedure RtfFooterEnd; virtual; abstract;
    procedure RtfTitle(Title: string; LevelOffset: integer = 0;
      withNumbers: boolean = true; Bookmark: string = ''); virtual; // level-indentated
    function RtfBookMark(const Text, BookmarkName: string;
      bookmarkNormalized: boolean): string;
    function RtfBookMarkString(const Text, BookmarkName: string;
      bookmarkNormalized: boolean): string; virtual; abstract;
    function RtfLinkTo(const aBookName, aText: string): TProjectWriter;
    function RtfPageRefTo(aBookName: string; withLink: boolean): TProjectWriter;
    function RtfField(const FieldName: string): string; virtual; abstract;
    function RtfLinkToString(const aBookName, aText: string;
      bookmarkNormalized: boolean=false): string; virtual; abstract;
    function RtfHyperlinkString(const http,text: string): string; virtual; abstract;
    function RtfPageRefToString(aBookName: string; withLink: boolean;
      BookMarkAlreadyComputed: boolean=false; Sequence: Integer=0): string; virtual; abstract;
      // Title is put without numbers
    procedure RtfSubTitle(const Title: string); virtual; abstract;
    // change font size % DefFontSize
    function RtfFont(SizePercent: integer): TProjectWriter;
    function RtfFontString(SizePercent: integer): string; virtual; abstract;
    function RtfBig(const text: string): TProjectWriter; // \par + bold + 110% size + \par
      virtual; abstract;
    function RtfGoodSized(const Text: string): string; virtual; abstract;
    procedure SetInfo(const aTitle, aAuthor, aSubject, aManager, aCompany: string); virtual; abstract;
    procedure Clear; virtual;
    procedure SaveToWriter(aWriter: TProjectWriter);
    procedure MovePortion(sourcePos, destPos, aLen: integer);
    function Data: string;
    property Last: TLastRTF read fLast write SetLast;
    property Len: Integer read WR.Len;
    property ColsCount: integer read fColsCount;
    property Landscape: boolean read fLandscape write SetLandscape;
  end;

  TRTF = class(TProjectWriter)
  protected
    procedure SetLast(const Value: TLastRTF); override;
    procedure SetLandscape(const Value: boolean); override;
    procedure RtfKeywords(line: string; const KeyWords: array of string; aFontSize: integer=80); override;
    constructor InternalCreate; override;
  public
    constructor Create(const aLayout: TProjectLayout;
      aDefFontSizeInPts: integer; // in points
      aCodePage: integer = 1252;
      aDefLang: integer = $0409; // french is $040c (1036)
      aLandscape: boolean = false; CloseManualy: boolean = false;
      aTitleFlat: boolean = false; const aF0: string = 'Calibri';
      const aF1: string = 'Consolas'; aMaxTitleOutlineLevel: integer=5); override;
    procedure InitClose; override; // if CloseManualy was true
    procedure SaveToFile(Format: TSaveFormat; OldWordOpen: boolean); override;

    function AddRtfContent(const s: string): TProjectWriter; override;
    procedure RtfList(line: string); override;
    procedure RtfPage; override;
    function RtfLine: TProjectWriter; override;
    function RtfPar: TProjectWriter; override;
    function RtfImageString(const Image: string; // 'SAD-4.1-functions.png 1101x738 85%'
      const Caption: string; WriteBinary: boolean; perc: integer): string; override;
    function RtfImage(const Image: string; const Caption: string = '';
      WriteBinary: boolean = true; const RtfHead: string = '\li0\fi0\qc'): TProjectWriter; override;
    /// if line is displayed as code, add it and return true
    procedure RtfCols(const ColXPos: array of integer; FullWidth: integer;
      VertCentered, withBorder: boolean; const RowFormat: string = ''); override;
    procedure RtfColsHeader(const text: string); override;
    procedure RtfRow(const Text: array of string; lastRow: boolean=false); // after RtfCols
      override;
    function RtfColsEnd: TProjectWriter; override;
    procedure RtfEndSection; override;
    function RtfParDefault: TProjectWriter; override;
    procedure RtfHeaderBegin(aFontSize: integer); override;
    procedure RtfHeaderEnd; override;
    procedure RtfFooterBegin(aFontSize: integer); override;
    procedure RtfFooterEnd; override;
    // level-indentated
    procedure RtfTitle(Title: string; LevelOffset: integer = 0;
      withNumbers: boolean = true; Bookmark: string = ''); override;
    // returns bookmarkreal
    function RtfBookMarkString(const Text, BookmarkName: string;
      bookmarkNormalized: boolean): string; override;
    function RtfLinkToString(const aBookName, aText: string;
      bookmarkNormalized: boolean): string; override;
    function RtfField(const FieldName: string): string; override;
    function RtfHyperlinkString(const http,text: string): string; override;
    function RtfPageRefToString(aBookName: string; withLink: boolean;
      BookMarkAlreadyComputed: boolean; Sequence: integer): string; override;
    // Title is put without numbers
    procedure RtfSubTitle(const Title: string); override;
    // idem with string
    function RtfFontString(SizePercent: integer): string; override;
    // \par + bold + 110% size + \par
    function RtfBig(const text: string): TProjectWriter; override;
    function RtfGoodSized(const Text: string): string; override;
    procedure SetInfo(const aTitle, aAuthor, aSubject, aManager, aCompany: string); override;
  end;

  THtmlTag = (hBold, hItalic, hUnderline, hCode, hBR, hBRList, hNavy, hNavyItalic,
    hNbsp, hPre, hAHRef, hTable, hTD, hTR, hP, hTitle, hHighlight, hLT, hGT, hAMP,
    hUL, hLI, hH, hTBody, hTHead, hTH);
  THtmlTags = set of THtmlTag;
  THtmlTagsSet = array[boolean,THtmlTag] of AnsiString;

  THTML = class(TProjectWriter)
  protected
    Level: integer;
    Current: THtmlTags;
    InTable, HasLT: boolean;
    Stack: array[0..20] of THtmlTags; // stack to handle { }
    fColsAreHeader: boolean;
    fColsMD: TIntegerDynArray;
    fContent,fAuthor,fTitle: string;
    fSavedWriter: TStringWriter;
    procedure SetLast(const Value: TLastRTF); override;
    procedure RtfKeywords(line: string; const KeyWords: array of string; aFontSize: integer=80); override;
    procedure SetLandscape(const Value: boolean); override;
    procedure WriteAsHtml(P: PAnsiChar; W: PStringWriter);
    function ContentAsHtml(const text: string): string;
    procedure SetCurrent(W: PStringWriter);
    procedure OnError(msg: string; const args: array of const);
  public
    constructor Create(const aLayout: TProjectLayout;
      aDefFontSizeInPts: integer; // in points
      aCodePage: integer = 1252;
      aDefLang: integer = $0409; // french is $040c (1036)
      aLandscape: boolean = false; aCloseManualy: boolean = false;
      aTitleFlat: boolean = false; const aF0: string = 'Calibri';
      const aF1: string = 'Consolas'; aMaxTitleOutlineLevel: integer=5); override;
    function Clone: TProjectWriter; override;
    procedure InitClose; override; // if CloseManualy was true
    procedure SaveToFile(Format: TSaveFormat; OldWordOpen: boolean); override;
    procedure Clear; override;

    function AddRtfContent(const s: string): TProjectWriter; override;
    procedure RtfList(line: string); override;
    procedure RtfPage; override;
    function RtfLine: TProjectWriter; override;
    function RtfPar: TProjectWriter; override;
    function RtfImageString(const Image: string; // 'SAD-4.1-functions.png 1101x738 85%'
      const Caption: string; WriteBinary: boolean; perc: integer): string; override;
    function RtfImage(const Image: string; const Caption: string = '';
      WriteBinary: boolean = true; const RtfHead: string = '\li0\fi0\qc'): TProjectWriter; override;
    /// if line is displayed as code, add it and return true
    procedure RtfCols(const ColXPos: array of integer; FullWidth: integer;
      VertCentered, withBorder: boolean; const RowFormat: string = ''); override;
    procedure RtfColsHeader(const text: string); override;
    procedure RtfRow(const Text: array of string; lastRow: boolean=false); override;
    function RtfColsEnd: TProjectWriter; override;
    procedure RtfEndSection; override;
    function RtfParDefault: TProjectWriter; override;
    procedure RtfHeaderBegin(aFontSize: integer); override;
    procedure RtfHeaderEnd; override;
    procedure RtfFooterBegin(aFontSize: integer); override;
    procedure RtfFooterEnd; override;
    procedure RtfTitle(Title: string; LevelOffset: integer = 0;
      withNumbers: boolean = true; Bookmark: string = ''); override;
    function RtfBookMarkString(const Text, BookmarkName: string;
      bookmarkNormalized: boolean): string; override;
    function RtfLinkToString(const aBookName, aText: string;
      bookmarkNormalized: boolean): string; override;
    function RtfField(const FieldName: string): string; override;
    function RtfHyperlinkString(const http,text: string): string; override;
    function RtfPageRefToString(aBookName: string; withLink: boolean;
      BookMarkAlreadyComputed: boolean; Sequence: integer): string; override;
    procedure RtfSubTitle(const Title: string); override;
    function RtfFontString(SizePercent: integer): string; override;
    function RtfBig(const text: string): TProjectWriter; override;
    function RtfGoodSized(const Text: string): string; override;
    procedure SetInfo(const aTitle, aAuthor, aSubject, aManager, aCompany: string); override;
  end;

  THeapMemoryStream = class(TMemoryStream)
  // allocates memory from Delphi heap (FastMM4) and not windows.Global*()
  // and uses bigger growing size -> a lot faster
  protected
    function Realloc(var NewCapacity: Longint): Pointer; override;
  end;


function TrimLastPeriod(const s: string; FirstCharLower: boolean = false): string; // delete last '.'
function RtfBackSlash(const Text: string): string; // RtfBackSlash('C:\Dir\')='C:\\Dir\\'
function RtfBookMarkName(const Name: string): string; // bookmark name compatible
function MM2Inch(mm: integer): integer; // MM2Inch(210)=11905
function Hex32(const C: cardinal): string; // return the hex value of a cardinal
function BookMarkHash(const s: string): string;
function ImageSplit(Image: string; out aFileName, iWidth, iHeight: string;
  out w,h, percent, Ext: integer): boolean;

{$ifdef DIRECTEXPORTTOWORD}
function RtfToDoc(Format: TSaveFormat; RtfFileName: string; OldWordOpen: boolean): boolean; // RTF -> native DOC format
{$endif}


function IsKeyWord(const KeyWords: array of string; const aToken: String): Boolean;
// aToken must be already uppercase
// note that 'array of string' deals with Const - not TStringDynArray

function IsNumber(P: PAnsiChar): boolean;

const
  VALID_PICTURES_EXT: array[0..3] of string =
  ('.JPG','.JPEG','.PNG','.EMF');

  PASCALKEYWORDS: array[0..99] of string =
  ('ABSOLUTE', 'ABSTRACT', 'AND', 'ARRAY', 'AS', 'ASM', 'ASSEMBLER',
   'AUTOMATED', 'BEGIN', 'CASE', 'CDECL', 'CLASS', 'CONST', 'CONSTRUCTOR',
   'DEFAULT', 'DESTRUCTOR', 'DISPID', 'DISPINTERFACE', 'DIV', 'DO',
   'DOWNTO', 'DYNAMIC', 'ELSE', 'END', 'EXCEPT', 'EXPORT', 'EXPORTS',
   'EXTERNAL', 'FAR', 'FILE', 'FINALIZATION', 'FINALLY', 'FOR', 'FORWARD',
   'FUNCTION', 'GOTO', 'IF', 'IMPLEMENTATION', 'IN', 'INDEX', 'INHERITED',
   'INITIALIZATION', 'INLINE', 'INTERFACE', 'IS', 'LABEL', 'LIBRARY',
   'MESSAGE', 'MOD', 'NEAR', 'NIL', 'NODEFAULT', 'NOT', 'OBJECT',
   'OF', 'OR', 'OUT', 'OVERRIDE', 'PACKED', 'PASCAL', 'PRIVATE', 'PROCEDURE',
   'PROGRAM', 'PROPERTY', 'PROTECTED', 'PUBLIC', 'PUBLISHED', 'RAISE',
   'READ', 'READONLY', 'RECORD', 'REGISTER', 'REINTRODUCE', 'REPEAT', 'RESIDENT',
   'RESOURCESTRING', 'SAFECALL', 'SET', 'SHL', 'SHR', 'STDCALL', 'STORED',
   'STRING', 'STRINGRESOURCE', 'THEN', 'THREADVAR', 'TO', 'TRY', 'TYPE',
   'UNIT', 'UNTIL', 'USES', 'VAR', 'VARIANT', 'VIRTUAL', 'WHILE', 'WITH', 'WRITE',
   'WRITEONLY', 'XOR');

  DFMKEYWORDS: array[0..4] of string = (
    'END', 'FALSE', 'ITEM', 'OBJECT', 'TRUE');

  CKEYWORDS: array[0..47] of string = (
  'ASM', 'AUTO', 'BREAK', 'CASE', 'CATCH', 'CHAR', 'CLASS', 'CONST', 'CONTINUE',
  'DEFAULT', 'DELETE', 'DO', 'DOUBLE', 'ELSE', 'ENUM', 'EXTERN', 'FLOAT', 'FOR',
  'FRIEND', 'GOTO', 'IF', 'INLINE', 'INT', 'LONG', 'NEW', 'OPERATOR', 'PRIVATE',
  'PROTECTED', 'PUBLIC', 'REGISTER', 'RETURN', 'SHORT', 'SIGNED', 'SIZEOF',
  'STATIC', 'STRUCT', 'SWITCH', 'TEMPLATE', 'THIS', 'THROW', 'TRY', 'TYPEDEF',
  'UNION', 'UNSIGNED', 'VIRTUAL', 'VOID', 'VOLATILE', 'WHILE');

  CSHARPKEYWORDS : array[0..86] of string = (
  'ABSTRACT', 'AS', 'BASE', 'BOOL', 'BREAK', 'BY3', 'BYTE', 'CASE', 'CATCH', 'CHAR',
  'CHECKED', 'CLASS', 'CONST', 'CONTINUE', 'DECIMAL', 'DEFAULT', 'DELEGATE', 'DESCENDING',
  'DO', 'DOUBLE', 'ELSE', 'ENUM', 'EVENT', 'EXPLICIT', 'EXTERN', 'FALSE', 'FINALLY',
  'FIXED', 'FLOAT', 'FOR', 'FOREACH', 'FROM', 'GOTO', 'GROUP', 'IF', 'IMPLICIT',
  'IN', 'INT', 'INTERFACE', 'INTERNAL', 'INTO', 'IS', 'LOCK', 'LONG', 'NAMESPACE',
  'NEW', 'NULL', 'OBJECT', 'OPERATOR', 'ORDERBY', 'OUT', 'OVERRIDE', 'PARAMS',
  'PRIVATE', 'PROTECTED', 'PUBLIC', 'READONLY', 'REF', 'RETURN', 'SBYTE',
  'SEALED', 'SELECT', 'SHORT', 'SIZEOF', 'STACKALLOC', 'STATIC', 'STRING',
  'STRUCT', 'SWITCH', 'THIS', 'THROW', 'TRUE', 'TRY', 'TYPEOF', 'UINT', 'ULONG',
  'UNCHECKED', 'UNSAFE', 'USHORT', 'USING', 'VAR', 'VIRTUAL', 'VOID', 'VOLATILE',
  'WHERE', 'WHILE', 'YIELD');

  MODULA2KEYWORDS: array[0..68] of string = (
  'ABS', 'AND', 'ARRAY', 'BEGIN', 'BITSET', 'BOOLEAN', 'BY', 'CAP', 'CARDINAL',
  'CASE', 'CHAR', 'CHR', 'CONST', 'DEC', 'DEFINITION', 'DIV', 'DO', 'ELSE',
  'ELSIF', 'END', 'EXCL', 'EXIT', 'EXPORT', 'FALSE', 'FLOAT', 'FOR', 'FROM',
  'GOTO', 'HALT', 'HIGH', 'IF', 'IMPLEMENTATION', 'IMPORT', 'IN', 'INC', 'INCL',
  'INTEGER', 'LONGINT', 'LOOP', 'MAX', 'MIN', 'MOD', 'MODULE', 'NIL', 'NOT',
  'ODD', 'OF', 'OR', 'ORD', 'POINTER', 'PROC', 'PROCEDURE', 'REAL', 'RECORD',
  'RECORD', 'REPEAT', 'RETURN', 'SET', 'SIZE', 'THEN', 'TO', 'TRUE', 'TRUNC',
  'TYPE', 'UNTIL', 'VAL', 'VAR', 'WHILE', 'WITH');

  XMLKEYWORDS: array[0..0] of string = ('');

  RTFEndToken: set of AnsiChar = [#0..#254]-['A'..'Z','a'..'z','0'..'9','-'];

  HTML_TAGS: THtmlTagsSet = (
    ('<b>','<i>','<u>','<code>','<br>','<li>','<font color="navy">','<font color="navy"><i>',
      '&nbsp;','<pre>','<a href="%s">','<table>','<td>','<tr>','<p>','<h3>',
      '<span style="background-color:yellow;">','&lt;','&gt;','&amp;','<ul>','<li>',
      '<h%d>','<tbody><tr>','<thead><tr>','<th>'),
    ('</b>','</i>','</u>','</code>','','','</font>','</i></font>',
      '','</pre>','</a>','</table>','</td>','</tr>','</p>','</h3>'#13#10,'</span>',
      '','','','</ul>','</li>','</h%d>','</tr></tbody>','</tr></thead>','</th>'));

  // format() parameters: [QuotedContent,QuotedAuthor,EscapedPageTitle]
  CONTENT_HEADER =
    '<!DOCTYPE html>' + #13#10 +
    '<html lang="en">' + #13#10 +
      '<head>' + #13#10 +
        '<meta charset="utf-8">' + #13#10 +
        '<meta http-equiv="X-UA-Compatible" content="IE=edge">' + #13#10 +
        '<meta name="viewport" content="width=device-width, initial-scale=1">' + #13#10 +
        '<meta name="description" content=%s>' + #13#10 +
        '<meta name="author" content=%s>' + #13#10 +
        '<title>%s</title>' + #13#10 +
        '<link rel="stylesheet" href="%ssynproject.css">'#13#10+
      '</head>'#13#10 +
      '<body>';
  CONTENT_FOOTER =
      #13#10'</div></div></div>'#13#10'</body></html>';
  SIDEBAR_HEADER =
      #1'<div class="sidebar"><div class="sidebarwrapper">'#1;
  SIDEBAR_FOOTER =
     #1#13#10'</div></div>'#13#10+
     '<div class="document"><div class="documentwrapper">'+
     '<div class="bodywrapper"><div class="body">'#13#10#1;

procedure CSVValuesAddToStringList(const aCSV: string; List: TStrings); overload;
// add all values in aCSV into List[]

procedure CSVValuesAddToStringList(P: PChar; List: TStrings); overload;
// add all CSV values in P into List[]



implementation

uses
  Windows,
{$ifdef DIRECTEXPORTTOWORD}
  ActiveX, ComObj, Variants,
{$endif}
  ProjectDiff; // for TMemoryMap


function TrimLastPeriod(const s: string; FirstCharLower: boolean = false): string;
begin
  if s='' then
    result := '' else begin
    if s[length(s)]<>'.' then
      result := s else
      result := copy(s,1,length(s)-1);
    if FirstCharLower then
      result[1] := NormToLower[result[1]];
  end;
end;

function MM2Inch(mm: integer): integer;
// MM2Inch(210)=11905 : A4 width paper size
begin
  result := ((1440*1000)*mm) div 25400;
end;


{ THeapMemoryStream = faster TMemoryStream using FastMM4 heap, not windows.GlobalAlloc() }

const
  MemoryDelta = $8000; // 32kb growing size Must be a power of 2

function THeapMemoryStream.Realloc(var NewCapacity: Integer): Pointer;
// allocates memory from Delphi heap (FastMM4) and not windows.Global*()
// and uses bigger growing size -> a lot faster
var i: integer;
begin
  if (NewCapacity > 0) then begin
    i := Seek(0,soFromCurrent); // no direct access to fSize -> use Seek() trick
    if NewCapacity=Seek(0,soFromEnd) then begin // avoid ReallocMem() if just truncate
      result := Memory;
      Seek(i,soFromBeginning);
      exit;
    end;
    NewCapacity := (NewCapacity + (MemoryDelta - 1)) and not (MemoryDelta - 1);
    Seek(i,soFromBeginning);
  end;
  Result := Memory;
  if NewCapacity <> Capacity then begin
    if NewCapacity = 0 then begin
      FreeMem(Memory);
      Result := nil;
    end else begin
      if Capacity = 0 then
        GetMem(Result, NewCapacity) else
        ReallocMem(Result, NewCapacity);
      if Result = nil then raise EStreamError.Create('THeapMemoryStream');
    end;
  end;
end;


{ TStringWriter }

function TStringWriter.Init: PStringWriter;
begin
  fGrowBy := 2048; // first FastMM4 fast (small block), and then medium block size
  len := 0;
  fData := '';
  result := @self;
end;

function TStringWriter.Init(const Args: array of const): PStringWriter;
begin
  Init;
  result := AddArray(Args);
end;

function TStringWriter.Init(GrowBySize: integer): PStringWriter;
begin
  if GrowBySize<512 then
    GrowBySize := 512; // to avoid bug in AddShort(s1,s2)
  fGrowBy := GrowBySize;
  len := 0;
  fData := '';
  result := @self;
end;

procedure TStringWriter.SaveToStream(aStream: TStream);
begin
  if len>0 then
    aStream.Write(fData[1],len);
end;

procedure TStringWriter.SaveToWriter(const aWriter: TStringWriter);
begin
  aWriter.Write(fData[1],len);
end;

procedure TStringWriter.SetData(const Value: string);
begin
  len := length(Value);
  fdata := Value;
end;

function Max(const A, B: Integer): Integer;
begin
  if A > B then
    Result := A else
    Result := B;
end;

function TStringWriter.Add(const s: string): PStringWriter;
begin
  result := @self;
  sLen := length(s);
  if sLen=0 then exit;
  if len+sLen>length(fData) then
    SetLength(fData,length(fData)+Max(sLen,fGrowBy));
  move(s[1],fData[len+1],sLen);
  inc(len,sLen);
end;

function TStringWriter.AddPChar(p: pChar): PStringWriter;
begin
  result := @self;
  sLen := StrLen(p);
  if sLen=0 then exit;
  if len+sLen>length(fData) then
    SetLength(fData,length(fData)+Max(sLen,fGrowBy));
  move(p^,fData[len+1],sLen);
  inc(len,sLen);
end;

function TStringWriter.AddStringOfChar(ch: char; Count: integer): PStringWriter;
begin
  result := @self;
  if Count<=0 then exit;
  if len+Count>length(fData) then
    SetLength(fData,length(fData)+Max(Count,fGrowBy));
  fillchar(fData[len+1],Count,ord(ch));
  inc(len,Count);
end;

function TStringWriter.RtfBackSlash(Text: string; trimLastBackSlash: boolean = false): PStringWriter;
var i,j: integer;
begin
  result := @self;
  if Text='' then exit;
  if trimLastBackSlash and (Text[length(Text)]='\') then
    SetLength(Text,length(Text)-1);
  i := pos('\',Text);
  if i=0 then begin
    Add(Text);
    exit;
  end;
  j := 1;
  repeat
    AddCopy(Text,j,i);
    Add('\');
    j := i+1;
    i := posEx('\',Text,j);
  until i=0;
  AddCopy(Text,j,maxInt);
end;

function TStringWriter.AddCopy(const s: string; Index, Count: Integer): PStringWriter;
begin
  result := @self;
  sLen := length(s)+1;
  if Index>=sLen then
    exit;
  if cardinal(Index)+cardinal(Count)>cardinal(sLen) then // cardinal: Count can be = maxInt
    Count := sLen-Index;
  if Count<=0 then exit;
  if len+Count>length(fData) then
    SetLength(fData,length(fData)+Max(Count,fGrowBy));
  move(s[Index],fData[len+1],Count);
  inc(len,Count);
end;

function TStringWriter.Add(c: char): PStringWriter;
begin
  inc(len);
  result := @self;
  if (pointer(fData)=nil) or (len>pInteger(cardinal(fData)-4)^) then
//  if len>length(fData) then
    SetLength(fData,length(fData)+fGrowBy);
  fData[len] := c;
end;

function TStringWriter.Add(const s1, s2: string): PStringWriter;
var L1, l2: integer;
begin
  L1 := length(s1);
  L2 := length(s2);
  sLen := len+L1;
  if sLen+L2>length(fData) then
    SetLength(fData,length(fData)+Max(L1+L2,fGrowBy));
  move(s1[1],fData[len+1],L1);
  move(s2[1],fData[sLen+1],L2);
  len := sLen+L2;
  result := @self;
end;

function TStringWriter.Add(const s: array of string): PStringWriter;
var i: integer;
begin
  for i := 0 to high(s) do
    Add(s[i]);
  result := @self;
end;

function TStringWriter.AddShort(const s1, s2: shortstring): PStringWriter;
begin
  sLen := len+ord(s1[0]);
  if sLen+ord(s2[0])>length(fData) then
    SetLength(fData,length(fData)+fGrowBy); // fGrowBy is always >255+255
  move(s1[1],fData[len+1],ord(s1[0]));
  move(s2[1],fData[sLen+1],ord(s2[0]));
  len := sLen+ord(s2[0]);
  result := @self;
end;

function TStringWriter.Add(p: pChar; pLen: integer): PStringWriter;
begin
  result := @self;
  if pLen>0 then begin
    if (pointer(fData)=nil) or (len+pLen>pInteger(cardinal(fData)-4)^) then
      SetLength(fData,length(fData)+Max(pLen,fGrowby));
    move(p^,fData[len+1],pLen);
    inc(len,pLen);
  end;
end;

function TStringWriter.AddArray(const Args: array of const): PStringWriter;
// with XML standard for Float values
var i: integer;
    tmp: shortstring;
begin
  for i := 0 to high(Args) do
  with Args[i] do
  case VType of
    vtChar:       Add(VChar);
    vtWideChar:   Add(string(widestring(VWideChar)));
    vtString:     AddShort(VString^);
    vtAnsiString: Add(string(VAnsiString));
    vtWideString: Add(string(VWideString));
    vtPChar:      AddPChar(VPChar);
    vtInteger:    begin str(VInteger,tmp); AddShort(tmp); end;
    vtPointer:    AddHex32(cardinal(VPointer));
    vtInt64:      AddInt64(VInt64^);
    vtExtended:   Add(@tmp[0],FloatToText(@tmp[0],VExtended^,fvExtended,ffGeneral,18,0));
    vtCurrency:   Add(@tmp[0],FloatToText(@tmp[0],VCurrency^,fvCurrency,ffGeneral,18,0));
  end;
  result := @self;
end;

const
  hexChars: array[0..15] of Char = '0123456789ABCDEF';
  
function Hex32ToPChar(dest: pChar; aValue: cardinal): pChar;
// group by byte (2 hex chars at once): faster and easier to read
begin
  case aValue of
    $0..$FF: begin
    dest[1] := HexChars[aValue and $F];
    dest[0] := HexChars[aValue shr 4];
    result := dest+2;
    end;
    $100..$FFFF: begin
    dest[3] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[2] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[1] := HexChars[aValue and $F];
    dest[0] := HexChars[aValue shr 4];
    result := dest+4;
    end;
    $10000..$FFFFFF: begin
    dest[5] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[4] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[3] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[2] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[1] := HexChars[aValue and $F];
    dest[0] := HexChars[aValue shr 4];
    result := dest+6;
    end;
    else begin //$1000000..$FFFFFFFF: begin
    dest[7] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[6] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[5] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[4] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[3] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[2] := HexChars[aValue and $F]; aValue := aValue shr 4;
    dest[1] := HexChars[aValue and $F];
    dest[0] := HexChars[aValue shr 4];
    result := dest+8;
    end;
  end;
end;

function Hex32(const C: cardinal): string;
// return the hex value of a cardinal
var tmp: array[0..7] of char;
begin
  SetString(result,tmp,Hex32ToPChar(tmp,C)-tmp);
end;

function Hex64ToPChar(dest: pChar; aValue: Int64): pChar;
function Write8(dest: pChar; aValue: Cardinal): pChar;
begin
  dest[7] := HexChars[aValue and $F]; aValue := aValue shr 4;
  dest[6] := HexChars[aValue and $F]; aValue := aValue shr 4;
  dest[5] := HexChars[aValue and $F]; aValue := aValue shr 4;
  dest[4] := HexChars[aValue and $F]; aValue := aValue shr 4;
  dest[3] := HexChars[aValue and $F]; aValue := aValue shr 4;
  dest[2] := HexChars[aValue and $F]; aValue := aValue shr 4;
  dest[1] := HexChars[aValue and $F];
  dest[0] := HexChars[aValue shr 4];
  result := dest+8;
end;
begin
  with Int64Rec(aValue) do
  if Hi<>0 then
    result := Write8(Hex32ToPChar(dest,Hi),Lo) else
    result := Hex32ToPChar(dest,Lo);
end;

function TStringWriter.Add(const Format: string; const Args: array of const): PStringWriter;
// all standard commands are supported, but the most useful are optimized
var i,j,  c, n, L: integer;
    decim: shortstring;
begin
  result := @self;
  n := length(Args);
  L := length(Format);
  if (n=0) or (L=0) then exit;
  i := 1;
  c := 0;
  while (i<=L) do begin
    j := i;
    while (i<=L) and (Format[i]<>'%') do inc(i);
    case i-j of
      0: ;
      1:   Add(Format[j]);
      else AddCopy(Format,j,i-j);
    end;
    inc(i);
    if i>L then break;
    if (Format[i] in ['0'..'9']) and (i<L) and (Format[i+1]=':') then begin
      c := ord(Format[i])-48;  // Format('%d %d %d %0:d %d',[1,2,3,4]) = '1 2 3 1 2'
      inc(i,2);
      if i>L then break;
    end;
    if Format[i]='%' then         // Format('%%') = '%'
      Add('%') else  // Format('%.3d',[4]) = '004':
    if (Format[i]='.') and (i+2<=L) and (c<n) and (Format[i+1] in ['1'..'9'])
      and (Format[i+2] in ['d','x']) and (Args[c].VType=vtInteger) then begin
      if Format[i+2]='d' then
        str(Args[c].VInteger,decim) else
        decim[0] := chr(Hex32ToPChar(@decim[1],Args[c].VInteger)-@decim[1]);
      for j := length(decim) to ord(Format[i+1])-49 do
        Add('0');
      AddShort(decim);
      inc(c);
      inc(i,2);
    end else
    if c<n then begin
      with Args[c] do
      case Format[i] of
      's': case VType of
        vtString:     AddShort(VString^);
        vtAnsiString: Add(string(VAnsiString));
        vtWideString: Add(string(VWideString));
        vtPChar:      AddPChar(VPChar);
        vtChar:       Add(VChar);
        vtInteger:    AddCardinal(VInteger); // extension from std FormatBuf
      end;
      'd': case VType of
        vtInteger: AddInteger(VInteger);
        vtInt64:   AddInt64(VInt64^);
      end;
      'x': case VType of
        vtInteger: AddHex32(VInteger);
        vtInt64:   AddHex64(VInt64^);
      end;
      else begin // all other formats: use standard FormatBuf()
        j := i-1; // Format[j] -> '%'
        while (i<=L) and not(NormToUpper[Format[i]] in ['A'..'Z']) do
          inc(i); // Format[i] -> 'g'
        Add(@decim[0],FormatBuf(decim[0],255,Format[j],i-j+1,Args[c]));
      end;
      end;
      inc(c);
    end;
    inc(i);
  end;
end;

function TStringWriter.AddWithoutPeriod(const s: string; FirstCharLower: boolean = false): PStringWriter;
var i: integer;
begin
  result := @self;
  if s='' then exit;
  slen := length(s);
  if s[slen]='.' then
   dec(slen);
  Add(pointer(s),slen);
  if FirstCharLower then begin
    i := len-slen;
    fData[i+1] := NormToLower[fData[i+1]];
  end;
end;

function TStringWriter.AddShort(const s: array of shortstring): PStringWriter;
var i: integer;
begin
  for i := 0 to high(s) do
    AddShort(s[i]);
  result := @self;
end;

function TStringWriter.DeleteLast(const s: string): boolean;
begin
  result := IsLast(s);
  if result then
    dec(len,length(s));
end;

function TStringWriter.IsLast(const s: string): boolean;
var L: integer;
begin
  L := length(s);
  result := (L=0) or ((L<=len) and (copy(fData,len-L+1,L)=s));
end;

function TStringWriter.EnsureLast(const s: string): PStringWriter;
// if last data is not s -> add(s)
begin
  result := @self;
  if not IsLast(s) then
    Add(s);
end;

function TStringWriter.RtfValid: boolean;
var i, Len1,Len2: integer;
begin
  Len1 := 0;
  Len2 := 0;
  for i := 1 to len do
    case fData[i] of
    '{': if (i=1) or (fData[i-1]<>'\') then inc(Len1);
    '}': if (i=1) or (fData[i-1]<>'\') then inc(Len2);
    end;
  result := Len1=Len2;
end;

function TStringWriter.AddShort(const s: shortstring): PStringWriter;
begin
  result := @self;
  sLen := ord(s[0]);
  if sLen=0 then exit;
  if (pointer(fData)=nil) or (len+sLen>pInteger(cardinal(fData)-4)^) then
//  if len+sLen>length(fData) then
    SetLength(fData,length(fData)+fGrowBy); // fGrowBy is always >255
  move(s[1],fData[len+1],sLen);
  inc(len,sLen);
end;

function TStringWriter.GetData: string;
begin
  if len<>length(fData) then
    SetLength(fData,len);
  result := fData;
end;

function TStringWriter.GetDataPointer: pointer;
begin
  result := pointer(fData);
end;

function TStringWriter.AddCardinal(value: cardinal): PStringWriter;
var tmp: string[15];
begin
  str(value,tmp);
  result := AddShort(tmp);
end;

function TStringWriter.AddInt64(value: Int64): PStringWriter;
// with FastMM4: using Int64Str (JOH version) is 8x faster than str()
begin
  result := Add(IntToStr(value));
end;

function TStringWriter.AddHex32(value: cardinal): PStringWriter;
var tmp: array[0..7] of char;
begin
  result := Add(tmp,Hex32ToPChar(tmp,value)-@tmp[0]);
end;

function TStringWriter.AddHex64(value: Int64): PStringWriter;
var tmp: array[0..15] of char;
begin
  result := Add(tmp,Hex64ToPChar(tmp,value)-@tmp[0]);
end;

function TStringWriter.AddInteger(const value: integer): PStringWriter;
begin
  result := Add(IntToStr(value));
end;

procedure TStringWriter.AddShortNoGrow(const s: shortstring);
var n: integer;    // faster: no Grow test
begin
  n := ord(s[0]);
  move(s[1],fData[len+1],n);
  inc(len,n);
end;

function TStringWriter.AddCRLF: PStringWriter;
begin
  result := @self;
  if (pointer(fData)=nil) or (len+2>pInteger(cardinal(fData)-4)^) then
    SetLength(fData,length(fData)+fGrowBy); // fGrowBy is always >2
  pWord(@fData[len+1])^ := $0a0d;  // CRLF = #13#10
  inc(len,2);
end;


procedure StringWriterTest;
var s: string;
    c: TStringWriter; // automatic Garbage Collector
begin
  fillchar(c,sizeof(c),0);
  c.Init.Add('<').Add(HexChars,4).Add('>');
  s := c.Add('Essai %s numéro %g et',['d''un',3.14159]).
    AddShort('</').AddByte(100).Add('>').Data;
  assert(s='<0123>Essai d''un numéro 3'+DecimalSeparator+'14159 et</100>');
end;

function ModDiv32(const Dividend, Divisor: Cardinal; out Quotient: Cardinal): Cardinal;
{ Quotient := Dividend mod Divisor;
  Result :=   Dividend div Divisor; }
asm
  push ecx
  mov ecx,edx
  xor edx,edx
  div ecx
  pop ecx
  mov [ecx],edx
end;

function TStringWriter.AddByte(value: byte): PStringWriter;
var dest: array[0..3] of char;
    tmp: cardinal;
begin
  if value<10 then
    result := Add(chr(value+48)) else
    if value<100 then // 10..99
      result := Add(@TwoDigitLookup[value],2) else begin // 100..999
      dest[0] := chr(ModDiv32(value,100,tmp)+48); // tmp=value mod 100
      pWord(@dest[1])^ := pWord(@TwoDigitLookup[tmp])^;
      result := Add(dest,3);
    end;
end;

function TStringWriter.AddWord(value: word): PStringWriter;
var dest: array[0..4] of char;
    tmp, t: cardinal;
begin
  if value<10 then
    result := Add(chr(value+48)) else
    if value<100 then // 10..99
      result := Add(@TwoDigitLookup[value],2) else begin
      t := ModDiv32(value,100,tmp); // t=value/100 tmp=value mod 100
      if t<10 then begin // 100..999
        dest[0] := chr(t+48);
        pWord(@dest[1])^ := pWord(@TwoDigitLookup[tmp])^;
        result := Add(dest,3);
      end else
      if t<100 then begin // 1000..9999
        pWord(@dest[0])^ := pWord(@TwoDigitLookup[t])^;
        pWord(@dest[2])^ := pWord(@TwoDigitLookup[tmp])^;
        result := Add(dest,4);
      end else begin // 10000..99999
        dest[0] := chr(ModDiv32(t,100,t)+48); // t=(value/100) mod 100
        pWord(@dest[1])^ := pWord(@TwoDigitLookup[t])^;
        pWord(@dest[3])^ := pWord(@TwoDigitLookup[tmp])^;
        result := Add(dest,5);
      end;
  end;
end;

function TStringWriter.AddCardinal6(cmd: char; value: cardinal; withLineNumber: boolean): PStringWriter;
// AddCardinal6('+',1234)=Add('+ 001234 ')
type PInt6 = ^TInt6;
     TInt6 = packed record cm,W0,W1,W2: TTwoDigit; sp: char; end;
var t,tmp: cardinal;
begin
  result := @self;
  if (pointer(fData)=nil) or (len+sizeof(TInt6)>pInteger(cardinal(fData)-4)^) then
    SetLength(fData,length(fData)+fGrowby);
  with PInt6(@fData[len+1])^ do begin
    cm[1] := cmd;
    cm[2] := ' ';
    inc(len,2);
    if not withLineNumber then
      exit;
    t := ModDiv32(value,100,tmp); // t=value/100 tmp=value mod 100
    W0 := TwoDigitLookup[ModDiv32(t,100,t)];
    W1 := TwoDigitLookup[t];
    W2 := TwoDigitLookup[tmp];
    sp := ' ';
  end;
  inc(len,sizeof(TInt6)-2);
end;

procedure TStringWriter.Write(const buf; bufLen: integer);
begin
  if bufLen<=0 then exit;
  if (pointer(fData)=nil) or (len+bufLen>pInteger(cardinal(fData)-4)^) then
    SetLength(fData,length(fData)+Max(bufLen,fGrowby));
  move(buf,fData[len+1],bufLen);
  inc(len,bufLen);
end;

procedure TStringWriter.Write(Value: integer);
begin
  if (pointer(fData)=nil) or (len+4>pInteger(cardinal(fData)-4)^) then
    SetLength(fData,length(fData)+fGrowby);
  move(Value,fData[len+1],4);
  inc(len,4);
end;

procedure TStringWriter.Write(const s: string);
var P: PChar;
    bufLen: integer;
begin
  P := pointer(S);
  if P=nil then exit;
  bufLen := PInteger(@P[-4])^;
  if (pointer(fData)=nil) or (len+bufLen>pInteger(cardinal(fData)-4)^) then
    SetLength(fData,length(fData)+Max(bufLen,fGrowby));
  move(P^,fData[len+1],bufLen);
  inc(len,bufLen);
end;

{function RtfWrite(Src,Dst: PChar): PChar;
var I: integer;
begin
  if Src<>nil then
  repeat
    if Src^=#0 then break else
    if Src^='\' then begin
      pWord(Dst)^ := ord('\')+ord('\')shl 8;
      inc(Dst,2);
      inc(Src);
    end else
    if Src^<#128 then begin
      Dst^ := Src^;
      inc(Dst);
      inc(Src);
    end else begin
      i := ord(Src^);
      Dst[0] := '\';
      Dst[1] := '''';
      Dst[2] := hexChars[i shr 4];
      Dst[3] := hexChars[i and $F];
      inc(Dst,4);
      inc(Src);
    end;
  until false;
  result := Dst;
end;

function TRTF.Rtf(const s: string): TProjectWriter;
begin
  result := self;
  WR.sLen := length(s)*4; // max chars that may be added
  if WR.sLen=0 then exit;
  if WR.len+WR.sLen>length(WR.fData) then
    SetLength(WR.fData,length(WR.fData)+Max(WR.sLen,WR.fGrowBy));
  WR.len := RtfWrite(pointer(s),PChar(pointer(WR.fData))+WR.len)-PChar(pointer(WR.fData));
end;}

function ImageSplit(Image: string; out aFileName, iWidth, iHeight: string;
  out w,h, percent, Ext: integer): boolean;
var i,j,err: integer;
begin
  result := false;
  if Image='' then exit;
  h := 0;
  w := 0;
  if Image[1]='%' then delete(Image,1,1);
  percent := 100;
  j := length(Image);
  if Image[j]='%' then // image width in page width %
    repeat
      dec(j);
      if not (Image[j] in ['0'..'9']) then begin
        percent := StrToIntDef(copy(Image,j+1,length(Image)-j-1),100);
        SetLength(Image,j-1);
        break;
      end;
    until false;
  aFileName := '';
  for i := length(Image) downto 1 do
    if Image[i]=' ' then begin
      aFileName := copy(Image,1,i-1);
      delete(Image,1,i);
      j := pos('x',Image);
      if j<2 then exit;
      iWidth := copy(Image,1,j-1);
      val(iWidth,w,err);
      if err<>0 then exit;
      iHeight := copy(Image,j+1,10);
      val(iHeight,h,err);
      if err<>0 then exit;
      break;
    end;
  j := 0;
  for i := length(aFileName) downto 1 do
    if aFileName[i]='.' then begin
      j := i;
      break;
    end;
  if j=0 then exit;
  Ext := GetStringIndex(VALID_PICTURES_EXT,copy(aFileName,j,5));
  if Ext>=0 then
    result := true;
end;

function TStringWriter.AddHexBuffer(value: PByte; Count: integer): PStringWriter;
var Dst: PChar;
    i: integer;
begin
  result := @self;
  sLen := Count*2; // chars count that are added
  if sLen=0 then exit;
  if len+sLen>length(fData) then
    SetLength(fData,length(fData)+Max(sLen,fGrowBy));
  Dst := pointer(fData);
  inc(Dst,len);
  inc(len,sLen);
  for i := 1 to Count do begin
    Dst[0] := hexChars[Value^ shr 4];
    Dst[1] := hexChars[Value^ and $F];
    inc(Value);
    inc(Dst,2);
  end;
end;

procedure TStringWriter.SaveToFile(const aFileName: String);
var F: TFileStream;
begin
  F := TFileStream.Create(aFileName,fmCreate);
  try
    SaveToStream(F);
  finally
    F.Free;
  end;
end;


function TStringWriter.GetPortion(aPos, aLen: integer): string;
begin
  result := Copy(fData,aPos,aLen);
end;

procedure TStringWriter.MovePortion(sourcePos, destPos, aLen: integer);
var Part: string;
begin
  Part := GetPortion(sourcePos, aLen);
  Insert(Part,fData,destPos);
  if sourcePos>destPos then
    inc(sourcePos,aLen);
  delete(fData,sourcePos,aLen);
end;


{ TProjectWriter }

constructor TProjectWriter.Create(const aLayout: TProjectLayout;
  aDefFontSizeInPts, aCodePage, aDefLang: integer; aLandscape,
  aCloseManualy, aTitleFlat: boolean; const aF0, aF1: string;
  aMaxTitleOutlineLevel: integer);
begin
  InternalCreate;
  aDefFontSizeInPts := aDefFontSizeInPts*2;
  FontSize := aDefFontSizeInPts;
  fLandscape := aLandscape;
  Layout := aLayout;
  if LandScape then begin
    Layout.Page.Width := aLayout.Page.Height;
    Layout.Page.Height := aLayout.Page.Width;
  end;
  Width := Layout.Page.Width-Layout.Margin.Left-Layout.Margin.Right;
  TitleFlat := aTitleFlat;
  TitleWidth := 0;     // title indetation gap (default 0 twips)
  IndentWidth := 240;  // list or title first line indent (default 240 twips)
  if not TitleFlat then
    TitleLevel[0] := -1; // introduction = first chapter = no number
  MaxTitleOutlineLevel := aMaxTitleOutlineLevel;
  // inherited should finish with:  if not CloseManualy then InitClose;
end;

procedure TProjectWriter.Clear;
begin
  WR.len := 0;
end;

procedure TProjectWriter.SaveToWriter(aWriter: TProjectWriter);
begin
  WR.SaveToWriter(aWriter.WR);
end;

function TProjectWriter.Clone: TProjectWriter;
begin
  result := TProjectWriterClass(ClassType).InternalCreate;
  result.FontSize := FontSize;
  result.Layout := Layout;
  result.Width := Width;
  result.PicturePath := PicturePath;
end;

constructor TProjectWriter.InternalCreate;
begin
  WR.Init;
end;

function TProjectWriter.AddRtfContent(const fmt: string;
  const params: array of const): TProjectWriter;
begin
  result := AddRtfContent(format(fmt,params));
end;

function TProjectWriter.AddWithoutPeriod(const s: string;
  FirstCharLower: boolean): TProjectWriter;
var tmp: string;
    L: integer;
begin
  result := self;
  if s='' then
    exit;
  L := length(s);
  if s[L]<>'.' then
    tmp := s else
    if L=1 then
      exit else
      SetString(tmp,PChar(pointer(s)),L-1);
  if FirstCharLower then
    tmp[1] := NormToLower[tmp[1]] else
    tmp[1] := NormToUpper[tmp[1]];
  AddRtfContent(tmp);
end;

procedure TProjectWriter.SetLandscape(const Value: boolean);
begin
  fLandscape := Value;
end;

procedure TProjectWriter.MovePortion(sourcePos, destPos, aLen: integer);
begin
  WR.MovePortion(sourcePos+1,destPos+1,aLen);
end;

function TProjectWriter.Data: string;
begin
  result := WR.Data;
end;

function TProjectWriter.RtfText: TProjectWriter;
begin
  SetLast(lastRtfText); // set Last -> update any pending {} (LastRtfList e.g.)
  fKeyWordsComment := false; // force end of comments
  result := self;
end;

procedure TProjectWriter.RtfColVertAlign(ColIndex: Integer;
  Align: TAlignment; DrawBottomLine: boolean);
var i, j: integer;
begin
  j := 1;
  for i := 0 to ColIndex do begin
   j := PosEx('\clvertal',fCols,j)+9;
   if j=9 then exit;
  end;
  case Align of // left=top, middle=center, right=bottom
    taLeftJustify:  fCols[j] := 't';
    taCenter:       fCols[j] := 'c';
    taRightJustify: fCols[j] := 'b';
  end;
  if DrawBottomLine then
    insert('\clbrdrb\brdrs\brdrcf9\brdrw24\brsp24',fCols,j-9) else
    if IdemPChar(@fCols[j-46],'\CLBRDR') then
      Delete(fCols,j-46,37);
end;

procedure TProjectWriter.RtfColsPercent(ColWidth: array of integer;
  VertCentered, withBorder, NormalIndent: boolean;
  RowFormat: string);
var x, w, i, ww: integer;
begin
  if NormalIndent then begin
    x := TitleWidth*pred(TitleLevelCurrent);
    ww := Width-x;
    if x<>0 then
      RowFormat := RowFormat+'\trleft'+IntToStr(x);
  end else begin
    x := 0;
    ww := Width;
  end;
  for i := 0 to high(ColWidth) do begin
    w := (ww*ColWidth[i]) div 100;
    inc(x,w);
    ColWidth[i] := x;
  end;
  RtfCols(ColWidth, ww, VertCentered, withBorder, RowFormat);
end;

function TProjectWriter.RtfCode(const line: string): boolean;
begin
  result := false;
  if line='' then
    exit;
  case line[1] of
    '!':
    case line[2] of
      '$': RtfDfm(copy(line,3,maxInt));
      else RtfPascal(copy(line,2,maxInt));
    end;
    '&': RtfC(copy(line,2,maxInt));
    '#': RtfCSharp(copy(line,2,maxInt));
    'µ': RtfModula2(copy(line,2,maxInt));
    '$': 
    case line[2] of
      '$': RtfSgml(copy(line,3,maxInt));
      else RtfListing(copy(line,2,maxInt));
    end;
  else exit;
  end;
  result := true;
end;

procedure TProjectWriter.RtfPascal(const line: string; aFontSize: integer = 80);
begin
  RtfKeywords(line, PASCALKEYWORDS, aFontSize);
end;

procedure TProjectWriter.RtfDfm(const line: string; aFontSize: integer);
begin
  RtfKeywords(line, DFMKEYWORDS, aFontSize);
end;

procedure TProjectWriter.RtfC(const line: string);
begin
  RtfKeywords(line, CKEYWORDS);
end;

procedure TProjectWriter.RtfCSharp(const line: string);
begin
  RtfKeywords(line, CSHARPKEYWORDS);
end;

procedure TProjectWriter.RtfModula2(const line: string);
begin
  RtfKeywords(line, MODULA2KEYWORDS);
end;

procedure TProjectWriter.RtfListing(const line: string);
begin
  RtfKeywords(line, []);
end;

procedure TProjectWriter.RtfSgml(const line: string);
begin
  RtfKeywords(line, XMLKEYWORDS);
end;

procedure TProjectWriter.RtfColLine(const line: string);
var P: PChar;
    withBorder, normalIndent: boolean;
    n, v, err: integer;
    Col: string;
    Cols: TIntegerDynArray;
    Colss: TStringDynArray;
begin
  if line='' then exit;
  if line[1]='%' then begin
     // |%=-40%30%30  = each column size -:no border =:no indent
     if fColsCount<>0 then // close any pending (and forgotten) table
       RtfColsEnd;
     P := @line[2];
     if P^=#0 then exit else // |% -> just close Table
     if P^='=' then begin //  = : no indent
       normalIndent := true;
       inc(P);
     end else
       normalIndent := false;
     if P^='-' then begin  // - : no border
       withBorder := false;
       inc(P);
     end else
       withBorder := true;
     SetLength(Cols,20);
     n := 0;
     repeat
       Col := GetNextItem(P,'%');
       if Col='' then break;
       val(Col,v,err);
       if err<>0 then continue;
       ProjectCommons.AddInteger(Cols,n,v);
     until false;
     SetLength(Cols,n);
     RtfColsPercent(Cols,true,withBorder,normalIndent,'\trkeep');
  end else begin // |text col 1|text col 2|text col 3
     SetLength(Colss,20);
     n := 0;
     P := pointer(line);
     repeat
       Col := GetNextItem(P,'|');
       ProjectCommons.AddString(Colss,n,Col);
     until P=nil;
     if n=0 then exit;
     SetLength(Colss,n);
     RtfRow(Colss);
  end;
end;

function TProjectWriter.RtfLinkTo(const aBookName, aText: string): TProjectWriter;
begin
  fStringPlain := true;
  WR.Add(RtfLinkToString(aBookName,aText,false));
  fStringPlain := false;
  result := self;
end;

function TProjectWriter.RtfPageRefTo(aBookName: string;
  withLink: boolean): TProjectWriter;
begin
  fStringPlain := true;
  WR.Add(RtfPageRefToString(aBookName,withLink,false));
  fStringPlain := false;
  result := self;
end;

function TProjectWriter.RtfBookMark(const Text, BookmarkName: string;
  bookmarkNormalized: boolean): string;
// return real BookMarkName
begin
  if BookmarkName='' then begin
    result := '';
    AddRtfContent(Text);
  end else begin
    if bookmarkNormalized then
      result := BookmarkName else
      result := RtfBookMarkName(BookmarkName);
    fStringPlain := true;
    WR.Add(RtfBookMarkString(Text,result,true));
    fStringPlain := false;
  end;
end;

function TProjectWriter.RtfFont(SizePercent: integer): TProjectWriter;
// change font size % DefFontSize
begin
  fStringPlain := true;
  WR.Add(RtfFontString(SizePercent));
  fStringPlain := false;
  result := self;
end;

procedure TProjectWriter.RtfTitle(Title: string; LevelOffset: integer;
  withNumbers: boolean; Bookmark: string);
var i: Integer;
    Num: string;
begin
  if LevelOffset>0 then
    TitleLevelCurrent := LevelOffset else
    TitleLevelCurrent := 0;
  for i := 1 to length(Title) do
    if Title[i]<>' ' then begin
      inc(TitleLevelCurrent,i);
      break;
    end;
  if TitleLevelCurrent=0 then exit;
  // TitleLevelCurrent = 1..length(TitleLevel)
  fLastWasRtfPageInRtfTitle := fLastWasRtfPage;
  SetLast(lastRtfTitle);
  if TitleLevelCurrent>1 then
    Delete(Title,1,TitleLevelCurrent-1-LevelOffset);
  if TitleLevelCurrent>length(TitleLevel) then
    TitleLevelCurrent := length(TitleLevel);
  inc(TitleLevel[TitleLevelCurrent-1]);
  if not TitleFlat and (TitleLevel[0]=0) then // intro = first chapter = no number
    withNumbers := false;
  fillchar(TitleLevel[TitleLevelCurrent],length(TitleLevel)-TitleLevelCurrent,0);
  fBookmarkInRtfTitle := '';
  if Bookmark='' then begin
    if TitlesList<>nil then // force every title to have a bookmark
      fBookmarkInRtfTitle := 'TITLE_'+IntToStr(TitlesList.Count);
  end else
  if Bookmark<>'- 'then begin // Bookmark='-' -> force only title
    fBookmarkInRtfTitle := RtfBookMarkName(BookMark);
    LastTitleBookmark := fBookmarkInRtfTitle;
  end;
  fInRtfTitle := '';
  if withNumbers then begin
    for i := 1 to TitleLevelCurrent do
      fInRtfTitle := fInRtfTitle+IntToStr(TitleLevel[i-1])+'.';
    fInRtfTitle := fInRtfTitle+' ';
  end;
  fInRtfTitle := fInRtfTitle+title;
  if (TitlesList<>nil) and (fBookmarkInRtfTitle<>'-') then begin
    // '1.2.3 Title|BookMark',pointer(3)
    if not FullTitleInTableOfContent then begin
      i := pos('\line',Title);
      if i>0 then
        SetLength(Title,i-1);
    end;
    i := pos('{',Title);
    if i>0 then
      SetLength(Title,i-1);
    if Title<>'' then begin
      if withNumbers then begin
        Num := '';
        for i := 1 to TitleLevelCurrent do
          Num := Num+IntToStr(TitleLevel[i-1])+'.';
        Title := Num+' '+Title;
      end;
      TitlesList.AddObject(Title+'|'+fBookmarkInRtfTitle,pointer(TitleLevelCurrent));
    end;
  end;
end;             


{ TRTF }

constructor TRTF.Create(const aLayout: TProjectLayout;
  aDefFontSizeInPts, aCodePage, aDefLang: integer;
  aLandscape, CloseManualy, aTitleFlat: boolean; const aF0, aF1: string;
  aMaxTitleOutlineLevel: integer);
var i: integer;
begin
  inherited;
  WR.AddShort('{\rtf1\ansi\ansicpg').AddWord(aCodePage).AddShort('\deff0\deffs').
    AddInteger(aDefFontSizeInPts).AddShort('\deflang').AddInteger(aDefLang).
    AddShort('{\fonttbl{\f0\fswiss\fcharset0 ').Add(aF0).
    AddShort(';}{\f1\fmodern\fcharset0 ').Add(aF1).Add(';}}'+
      '{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;'+
      '\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;'+
      '\red255\green255\blue150;\red255\green255\blue255;\red22\green43\blue90;'+
      '\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;'+
      '\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;'+
      '\red235\green235\blue230;}'+ // full default colortbl (for \highlight)
      '\viewkind4\uc1\paperw').AddInteger(Layout.Page.Width).
      AddShort('\paperh').AddInteger(Layout.Page.Height).
    AddShort('\margl').AddInteger(Layout.Margin.Left).
    AddShort('\margr').AddInteger(Layout.Margin.Right).
    AddShort('\margt').AddInteger(Layout.Margin.Top).
    AddShort('\margb').AddInteger(Layout.Margin.Bottom).
    AddShort('{\stylesheet{ Normal;}');
  for i := 1 to MaxTitleOutlineLevel do
    WR.AddShort('{\s').AddInteger(i).AddShort('\outlinelevel').AddInteger(i-1).
       AddShort('\snext0 heading ').AddInteger(i).AddShort(';}');
  WR.Add('}');
  if LandScape then
    WR.AddShort('\landscape');
(*  if aHeaderTop<>0 then
    AddShort('\headery').AddInteger(aHeaderTop);
  if aFooterBottom<>0 then
    AddShort('\footery').AddInteger(aFooterBottom);
  if Header<>'' then
    AddShort('{\header ').Add(Header).Add('}'); *)
  if not CloseManualy then
    InitClose;
end;

constructor TRTF.InternalCreate;
begin
  inherited;
  HandlePages := true;
end;

procedure TRTF.InitClose;
begin
  WR.AddShort(#13'{');
  RtfParDefault;
  Last := lastRtfNone;
end;

function TRTF.RtfImageString(const Image: string;
  const Caption: string; WriteBinary: boolean; perc: integer): string;
// 'SAD-4.1-functions.png 1101x738 85%' (100% default width)
const PICEXTTORTF: array[0..high(VALID_PICTURES_EXT)] of string =
    ('jpeg','jpeg','png','emf');
var Map: TMemoryMap;
    w,h, percent, Ext: integer;
    aFileName, iWidth, iHeight, content: string;
begin
  result := '';
  if ImageSplit(Image, aFileName, iWidth, iHeight, w,h,percent,ext) then
  if Map.DoMap(aFileName) or Map.DoMap(PicturePath+aFileName) then
  try
    // write to
    Last := lastRtfImage;
    if perc<>0 then
      percent := perc;
    percent := (Width*percent) div 100;
    h := (h*percent) div w;
    if WriteBinary then begin
      SetString(content,PAnsiChar(Map.buf),Map._size);
      content := '\bin'+IntToStr(Map._size)+' '+content;
    end else begin
      SetLength(content,Map._size*2);
      Map.ToHex(pointer(content));
    end;
    result := '{\pict\'+PICEXTTORTF[ext]+'blip\picw'+iWidth+'\pich'+iHeight+
      '\picwgoal'+IntToStr(percent)+'\pichgoal'+IntToStr(h)+
      ' '+content+'}';
    if Caption<>'' then
      result := result+'\line '+Caption+'\par';
  finally
    Map.UnMap;
  end;
end;

function TRTF.RtfImage(const Image, Caption: string; WriteBinary: boolean;
  const RtfHead: string): TProjectWriter;
begin
  WR.Add('{').Add(RtfHead);
  WR.Add(RtfImageString(Image,Caption,WriteBinary,0));
  WR.AddShort('}'#13);
  result := self;
end;

procedure TRTF.RtfPage;
begin
  RtfText;
  WR.AddShort(#13'\page');
  RtfParDefault;
  fLastWasRtfPage := true;
end;

function TRTF.RtfPar: TProjectWriter;
begin
  RtfText;
  WR.AddShort('\par'#13);
  result := Self;
end;

function TRTF.RtfLine: TProjectWriter;
begin
  RtfText;
  WR.AddShort('\line'#13);
  result := self;
end;

function IsKeyWord(const KeyWords: array of string; const aToken: String): Boolean;
// aToken must be already uppercase
var First, Last, I, Compare: Integer;
begin
  First := Low(Keywords);
  Last := High(Keywords);
  Result := True;
  while First<=Last do begin
    I := (First + Last) shr 1;
    Compare := CompareStr(Keywords[I],aToken);
    if Compare=0 then
      Exit else
    if Compare<0  then
      First := I+1 else
      Last := I-1;
  end;
  Result := False;
end;

procedure TRTF.SaveToFile(Format: TSaveFormat; OldWordOpen: boolean);
begin
  RtfText;
  WR.AddShort('}}');
  if not (Format in [fDoc,fPdf]) then
    exit;
  WR.SaveToFile(FileName);
  if OldWordOpen then
    Format := fDoc;
{$ifdef DIRECTEXPORTTOWORD}
  if RtfToDoc(Format,FileName,OldWordOpen) then begin
    DeleteFile(pointer(FileName)); // delete .rtf file and open .doc
    if Format=fPdf then
      FileName := ChangeFileExt(FileName,'.pdf') else
      FileName := ChangeFileExt(FileName,'.doc');
  end;
{$endif}
end;

procedure TRTF.RtfCols(const ColXPos: array of integer; FullWidth: integer;
  VertCentered, withBorder: boolean; const RowFormat: string);
var i: integer;
    s: string;
begin
  if Last=lastRtfTitle then
    WR.AddShort('{\sb0\sa0\fs12\par}'); // increase gap between title and table
  if withBorder and (length(ColXPos)>0) and (Last<>LastRtfNone) then
    WR.AddShort('{\sb0\sa0\fs6\brdrb\brdrs\brdrcf9\ri').
      AddInteger(FullWidth-ColXPos[High(ColXPos)]).AddShort('\par}');
  Last := lastRtfCols;
  WR.AddShort(#13'{\li0\fi0\ql ');
  fColsCount := length(ColXPos);
  fCols := '\trowd\trgaph108';
  fCols := fCols+RowFormat;
  for i := 0 to fColsCount-1 do begin
    s := '\cellx'+IntToStr(ColXPos[i]);
    if VertCentered then
      s := '\clvertalc'+s;
    if withBorder then // border top+bottom only (more modern look)
      fCols := fCols+'\clbrdrt\brdrs\brdrcf9\clbrdrb\brdrs\brdrcf9\cf9'+s else
//      fCols := fCols+'\clbrdrt\brdrs\clbrdrl\brdrs\clbrdrb\brdrs\clbrdrr\brdrs'+s else
    fCols := fCols+s;
  end;
end;

procedure TRTF.RtfRow(const Text: array of string; lastRow: boolean);
var i: integer;
begin
  if length(Text)<>fColsCount then begin
    for i := 1 to length(Text) do // bad count -> direct write, without table
      WR.Add(Text[i]).Add(' ');
    WR.AddShort('\par');
  end else begin
    WR.Add(fCols);
    for i := 0 to fColsCount-1 do
      WR.AddShort('\intbl ').Add(Text[i]).AddShort('\cell');
    WR.AddShort('\row'#13);
  end;
  if lastRow then
    RtfColsEnd;
end;

function TRTF.RtfColsEnd: TProjectWriter;
begin
  result := self;
  if fColsCount=0 then exit;
  WR.AddShort('}'#13);
  Last := LastRtfNone;
  fColsCount := 0;
end;

function IsNumber(P: PAnsiChar): boolean;
begin
  if P=nil then begin
    result := false;
    exit;
  end;
  while P^ in ['0'..'9'] do
    inc(P);
  result := (P^=#0);
end;

procedure TRTF.RtfTitle(Title: string; LevelOffset: integer = 0;
  withNumbers: boolean = true; Bookmark: string = '');
// Title is indentated with level
begin
  inherited; // set TitleLevel*[] and fLastWasRtfPageInRtfTitle+fBookmarkInRtfTitle
  if not TitleFlat and (TitleLevelCurrent=1) then begin
    if not fLastWasRtfPageInRtfTitle then
      WR.AddShort('\page');
    WR.AddShort('{\sb2400\sa140\brdrb\brdrs\brdrcf9\brdrw60\brsp60\ql\b\cf9\shad');
    RtfFont(240);
  end else begin
    WR.AddShort('{\sb280\sa0\ql\b\cf9\fi-').AddInteger(IndentWidth).
    AddShort('\li').AddInteger(TitleWidth*pred(TitleLevelCurrent)+IndentWidth);
    if TitleFlat then
      RtfFont(Max(140-TitleLevelCurrent*10,100)) else
      RtfFont(Max(180-TitleLevelCurrent*20,100));
  end;
  if TitleLevelCurrent<=MaxTitleOutlineLevel then
    WR.AddShort('\s').AddInteger(TitleLevelCurrent).Add(' ');
  RtfBookMark(fInRtfTitle,fBookmarkInRtfTitle,true);
  WR.AddShort('\par}'#13);
  if TitleWidth<>0 then
    WR.AddShort('\li').AddInteger(TitleWidth*pred(TitleLevelCurrent)).Add(' ');
end;

function TRTF.RtfBookMarkString(const Text, BookmarkName: string;
  bookmarkNormalized: boolean): string;
begin
  if bookmarkNormalized then
    result := BookmarkName else
    result := RtfBookMarkName(BookmarkName);
  result :='{\*\bkmkstart '+result+'}'+Text+'{\*\bkmkend '+result+'}';
end;

procedure TRTF.RtfSubTitle(const Title: string);
// Title is put without numbers
begin
  RtfText.WR.AddShort('{\sb280\sa0\ql\fi-').AddInteger(IndentWidth);
  if TitleWidth<>0 then
    WR.AddShort('\li').AddInteger(TitleWidth*pred(TitleLevelCurrent)+IndentWidth);
  RtfFont(Max(160-TitleLevelCurrent*20,100));
  WR.Add(Title).AddShort('\par}'#13);
end;

function TRTF.RtfFontString(SizePercent: integer): string;
// change font size % DefFontSize - result in string
begin
  result := '\fs'+IntToStr((FontSize*SizePercent)div 100)+' ';
end;

procedure TRTF.SetLast(const Value: TLastRTF);
begin
  if ListLine then // true -> \line, not \par in RtfList()
    if (Value<>LastRtfList) and (fLast=LastRtfList) then
      WR.AddShort('\sa80\par}'#13);
  fLast := Value;
  if Value<>lastRtfNone then
    fLastWasRtfPage := false;
end;

procedure TRTF.RtfList(line: string);
var c: char;
begin
  if line='' then exit;
  if ListLine then // true -> \line, not \par in RtfList()
    if (fLast<>LastRtfList) then begin
      if WR.DeleteLast('\par'#13) then
        WR.AddShort('\sa40\par\sa80');
      WR.AddShort(#13'{\ql\sa0\sb0\fi-').AddInteger(IndentWidth).
      AddShort('\li').AddInteger(TitleWidth*TitleLevelCurrent+IndentWidth).
      Add(' ');
    end else begin
      WR.AddShort('\par'#13);
    end;
  Last := lastRtfList;
  c := line[1];
  repeat
    delete(line,1,1);
  until (line='') or (line[1]<>' ');
  if line='' then exit;
  WR.Add(c).AddShort('\tab ').Add(line);
  if not ListLine then // true -> \line, not \par in RtfList()
    WR.AddShort('\par'#13);
end;

function TRTF.RtfBig(const text: string): TProjectWriter;
// bold + 110% size + \par
begin
  RtfText.WR.Add('{\b\cf9');
  RtfFont(110);
  WR.Add(Text).AddShort('\par}');
  result := self;
end;

function TRTF.RtfGoodSized(const Text: string): string;
begin
  if (length(Text)>8) and not IdemPChar(pointer(Text),'DI-') then
    Result := '{'+RtfFontString(91)+Text+'}' else
    Result := Text; 
end;


{$ifdef DIRECTEXPORTTOWORD}
var
  CoInitialized: boolean = false;


function RtfToDoc(Format: TSaveFormat; RtfFileName: string; OldWordOpen: boolean): boolean;
// convert an RTF file to native DOC format, using installed Word exe
var vMSWord, vMSDoc: variant;
    ClassID: TCLSID;
    Unknown: IUnknown;
    Disp: IDispatch;
    Path, DocFileName, FN: string;
    i: integer;
Const wdOpenFormatRTF = 3;
      wdOpenFormatAuto = 0;
      wdExportFormatPDF = 17;
      msoEncodingAutoDetect = 50001;
begin
  result := false; // leave as RTF if conversion didn't succeed
{$WARN SYMBOL_PLATFORM OFF}
  if DebugHook<>0 then
    exit; // OLE automation usualy fails inside Delphi IDE
{$WARN SYMBOL_PLATFORM ON}
  // 1. init MSWord OLE communication
  if not CoInitialized then begin
    CoInitialize(nil);
    CoInitialized := true;
  end;
  try
    try
      ClassID := ProgIDToClassID('Word.Application');
      if not Succeeded(GetActiveObject(ClassID, nil, Unknown)) or // get active
         not Succeeded(Unknown.QueryInterface(IDispatch, Disp)) then
         if not Succeeded(CoCreateInstance(ClassID, nil, CLSCTX_INPROC_SERVER or
        CLSCTX_LOCAL_SERVER, IDispatch, Disp)) then // create msWord instance
          exit; // no msWord -> leave as RTF
      vMSWord := Disp;
      while vMSWord.Templates.Count=0 do // Office 2010 fix
        // http://stackoverflow.com/questions/5913665/word-2010-automation-goto-bookmark
        Sleep(200);
      // 2. remote control of msWord to convert from .rtf to .doc
      Path := ExtractFilePath(RtfFileName);
      if Path='' then
        Path := GetCurrentDir else
        RtfFileName := ExtractFileName(RtfFileName);
      vMSWord.ChangeFileOpenDirectory(Path);
      // close any opened version
      for i := vMSWord.Documents.Count downto 1 do begin
        vMSDoc := vMSWord.Documents.Item(i);
        FN := ExtractFileName(vMSDoc.Name);
        if ValAt(FN,0,'.')=ValAt(RtfFileName,0,'.') then
          vMSDoc.Close(0,EmptyParam,EmptyParam);
      end;
      // invisible 'open' + 'save as' = convert .rtf to .doc
      if OldWordOpen then
        vMSDoc := vMSWord.Documents.Open(
          RtfFileName, //FileName,
          false) else
        vMSDoc := vMSWord.Documents.Open(
          RtfFileName, //FileName,
          false, //ConfirmConversions,
          false, //ReadOnly,
          false, //AddToRecentFiles,
          '', //PasswordDocument,
          '', //PasswordTemplate,
          true, //Revert,
          '', //WritePasswordDocument,
          '', //WritePasswordTemplate,
          wdOpenFormatRTF, //Format,
          msoEncodingAutoDetect, //Encoding,
          false //Visible,
          //OpenConflictDocument,
          //OpenAndRepair,
          //DocumentDirection,
          //NoEncodingDialog
      );
      try
        if Format=fPdf then begin
          DocFileName := ChangeFileExt(RtfFileName, '.pdf');
          vMSDoc.SaveAs(DocFileName, wdExportFormatPDF);
        end else begin
          DocFileName := ChangeFileExt(RtfFileName, '.doc');
          vMSDoc.SaveAs(DocFileName, wdOpenFormatAuto);
        end;
        result := true;
      finally
        vMSDoc.Close(0,EmptyParam,EmptyParam);
      end;
    finally
      vMSWord := unassigned;
    end;
  except
    on Exception do
      MessageBox(0,'Error during the document creation.'#13#13+
        'Did''nt you miss a \ or a } in your text?',nil,MB_ICONERROR);
  end;
end;
{$endif}

// RtfBackSlash('C:\Dir\')='C:\\Dir\\'
function RtfBackSlash(const Text: string): string;
var i: integer;
begin
  result := Text;
  i := pos('\',Text);
  if i>0 then
  repeat
    insert('\',result,i); // result[i]='\\'
    i := posEx('\',result,i+2);
  until i=0;
end;

function BookMarkHash(const s: string): string;
begin
  result := Hex32(Adler32Asm(0,pointer(s),length(s)));
end;

var
  RtfBookMarkNameName,
  RtfBookMarkNameValue: string; // cache for RtfBookMarkName()

function RtfBookMarkName(const Name: string): string;
// bookmark name compatible with rtf (AlphaNumeric+'_')
var i: integer;
begin
  if Name=RtfBookMarkNameName then begin
    result := RtfBookMarkNameValue;
    exit;
  end;
  RtfBookMarkNameName := Name;
  result := UpperCase(Name);
  for i := 1 to length(result) do
    if not (result[i] in ['0'..'9','A'..'Z']) then
      result[i] := '_';
  if Length(result)>40 then
    result := copy(result,1,16)+'_TRUNC_'+BookMarkHash(Name);
  RtfBookMarkNameValue := result;
end;

function TRTF.RtfLinkToString(const aBookName, aText: string;
  bookmarkNormalized: boolean): string;
begin
  if bookmarkNormalized then
    result := aBookName else
    result := RtfBookMarkName(aBookName);
  result := '{\field{\*\fldinst HYPERLINK \\l "'+result+'"}{\fldrslt '+aText+'}}';
end;

function TRTF.RtfHyperlinkString(const http,text: string): string;
begin
  result := '{\field{\*\fldinst{HYPERLINK'#13'"'+http+'"'#13'}}{\fldrslt{\ul'#13;
  if text='' then
    result := result+http+'..'+'}}}' else
    result := result+text+'}}}';
end;

function TRTF.RtfField(const FieldName: string): string;
begin
  result := '{\field{\*\fldinst '+FieldName+'}}';
end;

function TRTF.RtfPageRefToString(aBookName: string; withLink: boolean;
  BookMarkAlreadyComputed: boolean; Sequence: integer): string;
begin
  if aBookName='' then begin
    result := '';
    exit;
  end;
  if not BookMarkAlreadyComputed then
    aBookName := RtfBookMarkName(aBookName);
  result := '{\field{\*\fldinst PAGEREF "'+aBookName+'"}}';
  if withLink then
    result := '{\field{\*\fldinst HYPERLINK \\l "'+aBookName+'"}{\fldrslt '+
      result+'}}';
end;

procedure CSVValuesAddToStringList(const aCSV: string; List: TStrings); overload;
// add all values in aCSV into List[]
var P: PChar;
    s: string;
begin
  P := pointer(aCSV);
  if List<>nil then
  while P<>nil do begin
   s := GetNextItem(P);
   if s<>'' then
     List.Add(s);
  end;
end;

procedure CSVValuesAddToStringList(P: PChar; List: TStrings); overload;
// add all CSV values in P into List[]
var s: string;
begin
  if List<>nil then
  while P<>nil do begin
   s := GetNextItem(P);
   if s<>'' then
     List.Add(s);
  end;
end;

procedure TRTF.RtfEndSection;
begin
  WR.AddShort('\sect}\endhere{');
end;

function TRTF.RtfParDefault: TProjectWriter;
begin
  WR.AddShort('\pard\plain\sb80\sa80\qj');
  if TitleWidth<>0 then
    WR.AddShort('\li').AddInteger(TitleWidth*pred(TitleLevelCurrent));
  result := RtfFont(100);
end;

procedure TRTF.RtfFooterBegin(aFontSize: integer);
begin
  WR.AddShort('{\footer\cf9 ');
  if aFontsize<>0 then
    RtfFont(aFontSize);
end;

procedure TRTF.RtfFooterEnd;
begin
  WR.Add('}');
  fLastWasRtfPage := true;
end;

procedure TRTF.RtfHeaderBegin(aFontSize: integer);
begin
  WR.AddShort('{\header\cf9 ');
  if aFontsize<>0 then
    RtfFont(aFontSize);
end;

procedure TRTF.RtfHeaderEnd;
begin
  WR.Add('}');
end;

function TRTF.AddRtfContent(const s: string): TProjectWriter;
begin
  WR.Add(s);
  result := self;
end;

procedure TRTF.SetInfo(const aTitle, aAuthor, aSubject, aManager, aCompany: string);
begin
  WR.AddShort('{\info{\title ').Add(aTitle).
     AddShort('}{\author ').Add(aAuthor).
     AddShort('}{\operator ').Add(aAuthor).
     AddShort('}{\subject ').Add(aSubject).
     AddShort('}{\doccomm Document auto-generated by SynProject http://synopse.info').
     AddShort('}{\*\manager ').Add(aManager).
     AddShort('}{\*\company ').Add(aCompany).
     AddShort('}}');
end;

procedure TRTF.SetLandscape(const Value: boolean);
begin
  if Value=fLandscape then
    exit;
  inherited;
  if Value then begin
    WR.AddShort('\sect}'+
      '\sectd\ltrsect\lndscpsxn\pgwsxn').AddInteger(Layout.Page.Height). // switch paper size
      AddShort('\pghsxn').AddInteger(Layout.Page.Width).
      AddShort('\marglsxn').AddInteger(Layout.Margin.Left). // don't change margins
      AddShort('\margrsxn').AddInteger(Layout.Margin.Right).
      AddShort('\margtsxn').AddInteger(Layout.Margin.Top).
      AddShort('\margbsxn').AddInteger(Layout.Margin.Bottom).
      AddShort('\endhere{');
    Width := Layout.Page.Height-Layout.Margin.Left-Layout.Margin.Right;
  end else begin
    WR.AddShort('\sect}'+
      '\sectd\ltrsect\pghsxn').AddInteger(Layout.Page.Height). // switch paper size
      AddShort('\pgwsxn').AddInteger(Layout.Page.Width).
      AddShort('\marglsxn').AddInteger(Layout.Margin.Left).
      AddShort('\margrsxn').AddInteger(Layout.Margin.Right).
      AddShort('\margtsxn').AddInteger(Layout.Margin.Top).
      AddShort('\margbsxn').AddInteger(Layout.Margin.Bottom).
      AddShort('\endhere{');
    Width := Layout.Page.Width-Layout.Margin.Left-Layout.Margin.Right;
  end;
end;

procedure TRTF.RtfKeywords(line: string; const KeyWords: array of string;
  aFontSize: integer = 80);
function GetLineContent(start,count: integer): string;
var k: Integer;
begin
  result := Copy(line,start,count);
  if @KeyWords=@PASCALKEYWORDS then // not already done before
    for k := length(result) downto 1 do
      if result[k] in ['{','}'] then
        insert('\',result,k);
end;
var s, comment: string;
    i,j,k, cf: integer;
    CString, PasString, yellow, navy: boolean;
function IsDfmHexa(p: PAnsiChar): boolean;
begin
  result := false;
  while (p^=' ') do inc(p);
  while p^<>#0 do
    if p^ in ['0'..'9','A'..'Z','}','\'] then
      inc(p) else
      exit;
  result := true;
end;
var IgnoreKeyWord: boolean;
begin
  if line='' then begin
    AddRtfContent('\line'#13);
    exit;
  end;
  CString := (@KeyWords=@MODULA2KEYWORDS) or (@KeyWords=@CKEYWORDS)
          or (@KeyWords=@CSHARPKEYWORDS);
  PasString := (@KeyWords=@PASCALKEYWORDS) or (@KeyWords=@DFMKEYWORDS);
  for i := length(line) downto 1 do
    case line[i] of
    '{','}': if (@KeyWords<>@PASCALKEYWORDS) then insert('\',line,i);
    '\': if not IdemPChar(@PChar(pointer(line))[i],'LINE') then // '\line' is allowed
      insert('\',line,i);
    end;
  IgnoreKeyWord := (length(KeyWords)=0) or (@KeyWords=@XMLKEYWORDS);
  navy := false;
  if @KeyWords=@DFMKEYWORDS then begin
    i := pos(' = ',line);
    if i>0 then begin
      insert('\cf9 ',line,i+3); // right side in navy color
    end else
    if IsDfmHexa(pointer(line)) then begin
      IgnoreKeyWord := true;
      navy := true;
    end;
  end;
  if Last<>lastRtfCode then begin
    Last := lastRtfCode;
    AddRtfContent('{\sa0\sb80\f1\cbpat16\ql');
  end else
    AddRtfContent('{\sa0\sb0\f1\cbpat16\ql');
  RtfFont(aFontSize);
  yellow := line[1]='!';
  if yellow then
    i := 2 else
    i := 1;
  // write leading spaces before \highlight
  while line[i]=' ' do begin
    AddRtfContent(' ');
    inc(i);
  end;
  yellow := line[1]='!';
  if yellow then
    AddRtfContent('\highlight7 ');
  if navy then
    AddRtfContent('\cf9 ');
  // C-style '//' comment detection
  comment := '';
  if not IgnoreKeyWord and not (@KeyWords=@DFMKEYWORDS) then begin
    j := Pos('//',line);
    if j>0 then begin
      comment := GetLineContent(j,maxInt);
      SetLength(line,j-1);
      if i>length(line) then
        line := '';
    end;
  end;
  if line<>'' then
    if @KeyWords=@XMLKEYWORDS then begin
      cf := 0;
      repeat
        case line[i] of
        #0: break;
        '<': begin AddRtfContent('\cf12 '); cf := 12; end;
        '>': begin if cf<>12 then AddRtfContent('\cf12 ');
                   AddRtfContent('>\cf0 '); cf := 0; inc(i); continue; end;
        '&': begin AddRtfContent('\cf10 '); cf := 10; end;
        ';': if cf=10 then begin
               AddRtfContent(';\cf0 '); cf := 0; inc(i); continue; end;
        '"': if cf=9 then begin
               AddRtfContent('"\cf12 '); cf := 12; inc(i); continue; end else
             if cf=12 then begin
               AddRtfContent('"\cf9 '); cf := 9; inc(i); continue; end;
        end;
        AddRtfContent(line[i]);
        inc(i);
      until false;
    end else
    if IgnoreKeyWord then // for '$ Simple listing' or DFM hexa
      AddRtfContent(Copy(line,i,maxInt)) else
    repeat
      // write leading spaces
      while line[i]=' ' do begin
        AddRtfContent(line[i]);
        inc(i);
      end;
      // '(* Pascal or Modula-2 comment *)'
      if (Line[i]='(') and (Line[i+1]='*') and
         ((@KeyWords=@PASCALKEYWORDS) or (@KeyWords=@MODULA2KEYWORDS)) then begin
        k := PosEx('*)',line,i+2);
        if k>0 then begin
          inc(k,2);
          AddRtfContent('{\cf9\i '+GetLineContent(i,k-i)+'}');
          i := k;
          continue;
        end;
      end;
      // '{ pascal s }'
      if @KeyWords=@PASCALKEYWORDS then
      case Line[i] of
      '{': begin
        k := PosEx('}',line,i+1);
        if k>0 then begin
          inc(k);
          AddRtfContent('{\cf9\i\{'+GetLineContent(i+1,k-i-2)+'\}}');
          i := k;
          continue;
        end else begin
          AddRtfContent('\{');
          inc(i);
        end;
      end;
      '}': begin // a line with no beginning of } -> full line of comment
          AddRtfContent('\}');
          inc(i);
      end;
      end;
      // '/* C comment */'
      if ((@KeyWords=@CKEYWORDS)or(@KeyWords=@CSHARPKEYWORDS)) and
        (Line[i]='/') and (Line[i+1]='*') then begin
        k := PosEx('*/',line,i+2);
        if k>0 then begin
          inc(k,2);
          AddRtfContent('{\cf9\i '+Copy(line,i,k-i)+'}');
          i := k;
          continue;
        end;
      end;
      // get next word till end word char
      j := 0;
      k := i;
      while line[k]<>#0 do
        if (@KeyWords=@MODULA2KEYWORDS) and (line[k]='|') then begin
          j := k;
          break;
        end else
        if CString and (line[k]='"') then begin
          if (line[k+1]='"') then
            // ignore ""
            inc(k,2) else begin
            j := k;
            break;
          end;
        end else
        if line[k]='''' then begin
          if PasString and (line[k+1]='''') then
            // ignore pascal ''
            inc(k,2) else begin
            j := k;
            break;
          end;
        end else
        if not (line[k] in ['0'..'9','a'..'z','A'..'Z','_']) then begin
          // we reached a non identifier-valid char, i.e. an end word char
          j := k;
          break;
        end else
          inc(k);
      // write word as keyword if the case
      if j=0 then
        s := copy(line,i,maxInt) else
        s := copy(line,i,j-i);
      if s<>'' then
        if (CString or PasString) and IsNumber(pointer(s)) then
          // write const number in blue
          AddRtfContent('{\cf9 '+s+'}') else
        if ((@KeyWords=@MODULA2KEYWORDS) and
              IsKeyWord(KeyWords,s)) or // .MOD keys are always uppercase
           ((@KeyWords<>@MODULA2KEYWORDS) and IsKeyWord(KeyWords,UpperCase(s))) then
          // write keyword in bold
          AddRtfContent('{\b '+s+'}') else
          // write normal word
          AddRtfContent(s);
      // handle end word char (in line[j])
      if j=0 then break;
      i := j+1;
      if (@KeyWords=@MODULA2KEYWORDS) and (line[j]='|') then
        AddRtfContent('{\b |}') else // put CASE | in bold
      if PasString and (line[j]='''') and (line[i]<>'''') then begin
        AddRtfContent('{\cf9 ''');
        while line[i]<>#0 do begin // put pascal 'string'
          if (@KeyWords=@PASCALKEYWORDS) and (line[i] in ['{','}']) then
            AddRtfContent('\');
          AddRtfContent(line[i]);
          inc(i);
          if line[i-1]='''' then break;
        end;
        AddRtfContent('}');     
        if line[i]=#0 then break;
      end else
      if CString and (line[j]='"') and (line[i]<>'"') then begin
        AddRtfContent('{\cf9 "');
        while line[i]<>#0 do begin // handle "string"
          AddRtfContent(line[i]);
          inc(i);
          if line[i-1]='"' then break;
        end;
        AddRtfContent('}');
        if line[i]=#0 then break;
      end else begin
        if (@KeyWords=@PASCALKEYWORDS) and (line[j] in ['{','}']) then
          AddRtfContent('\');
        AddRtfContent(line[j]); // write end word char
      end;
    until false;
  if comment<>'' then // write comment in italic
    AddRtfContent('{\cf9\i '+comment+'}');
  AddRtfContent('\par}'#13);
end;

procedure TRTF.RtfColsHeader(const text: string);
begin
  RtfColsPercent([100],true,false,false,'\trkeep');
  RtfRow(['\ql\line{\b\fs30 '+text+'}'],true);
end;


{ THTML }

function HtmlEncode(const s: string): string;
var i: integer;
begin // not very fast, but working
  result := '';
  for i := 1 to length(s) do
    case s[i] of
      '<': result := result+'&lt;';
      '>': result := result+'&gt;';
      '&': result := result+'&amp;';
      '"': result := result+'&quot;';
      else result := result+s[i];
    end;
end;

function THTML.AddRtfContent(const s: string): TProjectWriter;
begin
  WriteAsHtml(pointer(s),nil);
  result := self;
end;

constructor THTML.Create(const aLayout: TProjectLayout; aDefFontSizeInPts,
  aCodePage, aDefLang: integer; aLandscape, aCloseManualy,
  aTitleFlat: boolean; const aF0, aF1: string; aMaxTitleOutlineLevel: integer);
begin
  inherited;
  { TODO : multi-page layout }
end;

function THTML.Clone: TProjectWriter;
begin
  result := inherited Clone;
  (result as THTML).SetInfo(fTitle,fAuthor,fContent,'','');
end;

procedure THTML.SetInfo(const aTitle,aAuthor,aSubject,aManager,aCompany: string);
begin
  if aTitle<>'' then
    fTitle := aTitle;
  if aAuthor<>'' then
    fAuthor := aAuthor;
  if aSubject<>'' then
    fContent := aSubject;
end;

procedure THTML.Clear;
begin
  inherited Clear;
  Level := 0;
  fillchar(Stack,sizeof(Stack),0);
  Current := [];
end;

procedure THTML.InitClose;
begin
  inherited;
  { TODO : multi-page layout }
end;

var ConsoleAllocated: boolean;

procedure THTML.OnError(msg: string; const args: array of const);
begin
  if not ConsoleAllocated then
    AllocConsole;
  msg := Format(msg,args);
  writeln(msg);
//  MessageBox(0,pointer(msg),nil,MB_OK);
end;

function THTML.RtfBig(const text: string): TProjectWriter;
begin
  WR.Add('<h2>');
  AddRtfContent(text);
  WR.Add('</h2>').AddCRLF;
  result := self;
end;

function THTML.RtfBookMarkString(const Text, BookmarkName: string;
  bookmarkNormalized: boolean): string;
begin
  if BookmarkName='' then
    result := '' else begin
    if bookmarkNormalized then
      result := BookmarkName else
      result := RtfBookMarkName(BookmarkName);
    result := '<a name="'+result+'"></a>';
  end;
  if Text<>'' then
    if fStringPlain then
      result := result+ContentAsHtml(Text) else
      result := #1+result+#1+Text;
end;

procedure THTML.RtfCols(const ColXPos: array of integer;
  FullWidth: integer; VertCentered, withBorder: boolean;
  const RowFormat: string);
var left,i,w,md: integer;
begin
  fColsCount := length(ColXPos);
  if fColsCount=0 then
    exit;
  i := Pos('\trleft',RowFormat);
  if i>0 then
    left := StrToIntDef(copy(RowFormat,i+7,10),0) else
    left := 0;
  SetLength(fColsMD,fColsCount);
  for i := 0 to fColsCount-1 do begin
    w := ColXPos[i]-left;
    left := ColXPos[i];
    md := round((w*100)/Width);
    if md=0 then
      inc(md);
    fColsMD[i] := md;
  end;
  if Last=lastRtfCols then begin
    fColsAreHeader := false;
  end else begin
    SetLast(lastRtfCols);
    fColsAreHeader := IdemPChar(pointer(RowFormat),'\TRHDR');
    WR.AddCRLF.AddShort('<table class="docutils" align="center" width="100%"><colgroup>');
    for i := 0 to fColsCount-1 do
      WR.Add('<col width="%d%%" />',[fColsMD[i]]);
    WR.AddShort('</colgroup>');
  end;
end;

procedure THTML.RtfRow(const Text: array of string; lastRow: boolean);
var i: integer;
    t1,t2: THtmlTag;
    initial,save: THtmlTags;
begin
  if length(Text)<=fColsCount then begin
    if fColsAreHeader then begin
      t1 := hTHead;
      t2 := hTH;
    end else begin
      t1 := hTR;
      t2 := hTD;
    end;
    SetCurrent(@WR);
    initial := Current;
    save := initial;
    WR.Add(HTML_TAGS[false,t1]);
    for i := 0 to high(Text) do begin
      WR.Add(HTML_TAGS[false,t2]);
      Stack[Level] := save;
      SetCurrent(@WR);
      AddRtfContent(Text[i]);
      save := Stack[Level];
      Stack[Level] := initial;
      SetCurrent(@WR);
      WR.Add(HTML_TAGS[true,t2]);
    end;
    WR.Add(HTML_TAGS[true,t1]).AddCRLF;
    Stack[Level] := initial;
    SetCurrent(@WR);
  end;
  if lastRow then
    RtfColsEnd;
end;

function THTML.RtfColsEnd: TProjectWriter;
begin
  result := self;
  if fColsCount=0 then exit;
  if not fColsAreHeader then begin
    WR.Add(HTML_TAGS[True,hTable]).AddCRLF;
    Last := LastRtfNone;
  end;
  fColsCount := 0;
end;

procedure THTML.RtfColsHeader(const text: string);
begin
  WR.Add('<strong>').Add(text).Add('</strong>').AddCRLF;
end;

procedure THTML.RtfEndSection;
begin
end;

function THTML.RtfFontString(SizePercent: integer): string;
begin // any implementation should handle fStringPlain flag
  result := ''; // do nothing in responsive design
end;

procedure THTML.RtfFooterBegin(aFontSize: integer);
begin
  fSavedWriter := WR;
end;

procedure THTML.RtfFooterEnd;
begin
  WR := fSavedWriter;
end;

procedure THTML.RtfHeaderBegin(aFontSize: integer);
begin
  fSavedWriter := WR;
end;

procedure THTML.RtfHeaderEnd;
begin
  WR := fSavedWriter;
end;

function THTML.RtfGoodSized(const Text: string): string;
begin
  result := Text;
end;

function THTML.RtfImageString(const Image, Caption: string;
  WriteBinary: boolean; perc: integer): string;
// !!! TODO: produce SVG content instead of .emf binary
var w,h, percent, Ext, i,j: integer;
    aFileName, alt, src,iWidth, iHeight: string;
const EXTS: array[0..high(VALID_PICTURES_EXT)] of string = (
  'jpg','jpg','png','svg');
begin
  result := '';
  if not ImageSplit(Image, aFileName, iWidth, iHeight, w,h,percent,ext) then begin
    OnError('ImageSplit(%s)',[Image]);
    exit;
  end;
  if not FileExists(aFileName) then begin
    aFileName := PicturePath+aFileName;
    if not FileExists(aFileName) then begin
      OnError('File "%s" not found',[aFileName]);
      exit;
    end;
  end;
  if Ext=3 then begin // embed svg within the html document!
    aFileName := ChangeFileExt(aFileName,'.svg');
    FileToString(aFileName,src);
    assert(pos('''',src)=0);
    i := Pos('<g ',src);
    if i>1 then
      delete(src,1,i-1);
    src := StringReplaceAll(StringReplaceAll(src,'Times New Roman','Arial'),#13,'');
    repeat
      i := Pos('<g ',src);
      if i=0 then break;
      j := PosEx('>',src,i);
      if j<i then break;
      delete(src,i,j-i+1);
    until false;
    repeat
      i := Pos('<title>',src);
      if i=0 then break;
      j := PosEx('</title>',src,i);
      if j<i then break;
      delete(src,i,j-i+8);
    until false;
    src := StringReplaceAll(src,'</g>','');
    src := StringReplaceAll(src,'</svg>','</g></svg>');
    src := trim(StringReplaceAll(src,'&amp;','&'));
    result := format('<svg viewBox="0 0 %d %d" width="%d" height="%d">'+
      '<g id="graph" class="graph" style="font-size:18.67">',
      [w,h,(w*2)div 3,(h*2)div 3])+src;
  end else begin
    alt := ExtractFileName(aFileName);
    if DestPath='' then // ='' if from sub html
      result := '<img src="../img/'+alt+'" alt="'+alt+'">' else begin
      result := '<img src="img/'+alt+'" alt="'+alt+'">';
      alt := IncludeTrailingPathDelimiter(DestPath)+'img\'+alt;
      if not FileExists(alt) then begin
        if not DirectoryExists(ExtractFilePath(alt)) then
          CreateDir(ExtractFilePath(alt));
        CopyFile(aFileName,alt);
      end;
    end;
  end;
  if Caption<>'' then
    result := format('<figure>%s<figcaption>%s</figcaption></figure>',
      [result,ContentAsHtml(StringReplaceAll(Caption,'\line',''))]);
  if not fStringPlain then
    result := #1+result+#1;
end;

function THTML.RtfImage(const Image, Caption: string; WriteBinary: boolean;
  const RtfHead: string): TProjectWriter;
begin
  fStringPlain := true;
  WR.Add(RtfImageString(Image,Caption,WriteBinary,0));
  fStringPlain := false;
  result := self;
end;

function THTML.RtfLine: TProjectWriter;
begin
  RtfText;
  WR.Add(HTML_TAGS[false,hBR]).AddCRLF;
  result := self;
end;

function THTML.RtfLinkToString(const aBookName, aText: string;
  bookmarkNormalized: boolean): string;
var i: integer;
    f,b: string;
begin
  i := pos('#',aBookName);
  if i>0 then begin
    f := copy(aBookName,1,i-1);
    b := copy(aBookName,i+1,100);
  end else
    b := aBookName;
  if b<>'' then begin
    if not bookmarkNormalized then
      b := RtfBookMarkName(b);
    result := f+'#'+b;
  end else
    result := f;
  result := Format(HTML_TAGS[false,hAHref],[result])+ContentAsHtml(aText)+
    HTML_TAGS[true,hAHref];
  if not fStringPlain then
    result := #1+result+#1;
end;

function THTML.RtfHyperlinkString(const http,text: string): string;
begin
  result := Format(HTML_TAGS[false,hAHref],[http]);
  if text='' then
    result := result+http else
    result := result+ContentAsHtml(text);
  result := result+HTML_TAGS[true,hAHref];
  if not fStringPlain then
    result := #1+result+#1;
end;

function THTML.RtfField(const FieldName: string): string;
begin
  result := ''; // no field
end;

procedure THTML.SetLast(const Value: TLastRTF);
begin
  if Value=fLast then
    exit;
  Stack[Level] := [];
  if Current<>[] then
    SetCurrent(@WR);
  case fLast of
  lastRtfCode: WR.Add(HTML_TAGS[true,hPre]);
  lastRtfList: WR.Add(HTML_TAGS[true,hUL]);
  lastRtfText: WR.Add(HTML_TAGS[true,hP]).AddCRLF;
  end;
  case Value of
  lastRtfText: WR.Add(HTML_TAGS[false,hP]);
  end;
  fLast := Value;
end;

procedure THTML.RtfList(line: string);
begin
  while (line<>'') and (line[1] in [' ','-']) do
    delete(line,1,1);
  if line='' then
    exit;
  if Last<>lastRtfList then begin
    SetLast(lastRtfList);
    WR.Add(HTML_TAGS[false,hUL]);
  end;
  WR.Add(HTML_TAGS[false,hLI]);
  WR.Add(ContentAsHtml(line));
  WR.Add(HTML_TAGS[true,hLI]);
end;

function THTML.RtfPar: TProjectWriter;
begin
  SetLast(lastRtfNone);
  RtfText;
  result := RtfText;
end;

function THTML.RtfParDefault: TProjectWriter;
begin
  Stack[Level] := [];
  SetCurrent(@WR);
  RtfText;
  result := self;
end;

procedure THTML.RtfKeywords(line: string; const KeyWords: array of string;
  aFontSize: integer);
label com, str, str2;
var P,B: PAnsiChar;
    token: AnsiString;
    HasLT,HasGT,Highlight, CString, PasString, IgnoreKeyWord: boolean;
procedure SetToken;
begin
  SetString(token,B,P-B);
  if HasLT and (HTML_TAGS[false,hLT][1]<>'<') then
    token := StringReplaceAll(token,'<',HTML_TAGS[false,hLT]);
  if HasGT and (HTML_TAGS[false,hGT][1]<>'>') then
    token := StringReplaceAll(token,'>',HTML_TAGS[false,hGT]);
end;
begin
  if line='' then
    exit;
  if Last<>lastRtfCode then begin
    SetLast(lastRtfCode);
    WR.Add(HTML_TAGS[false,hPre]);
  end else
    WR.AddCRLF;
  P := Pointer(line);
  CString := (@KeyWords=@MODULA2KEYWORDS) or (@KeyWords=@CKEYWORDS)
          or (@KeyWords=@CSHARPKEYWORDS);
  PasString := (@KeyWords=@PASCALKEYWORDS) or (@KeyWords=@DFMKEYWORDS);
  IgnoreKeyWord := (length(KeyWords)=0) or (@KeyWords=@XMLKEYWORDS);
  Highlight := (P^='!');
  if Highlight then
    inc(P);
  B := P;
  while B^=' ' do inc(B);
  if B^<' ' then
    exit;
  if IgnoreKeyWord then begin
    if HighLight then
      WR.Add(HTML_TAGS[false,hHighlight]);
    HasLT := False;
    HasGT := False;
    B := P;
    while P^>=' ' do begin
      if P^='<' then HasLT := True else
      if P^='>' then HasGT := True;
      inc(P);
    end;
    SetToken;
    WR.Add(token);
    if HighLight then
      WR.Add(HTML_TAGS[true,hHighlight]);
    exit;
  end;
  while P^=' ' do begin
    WR.Add(' '); // HTML_TAGS[false,hNbsp]); needed only for blog 
    inc(P);
  end;
  if HighLight then
    WR.Add(HTML_TAGS[false,hHighlight]);
  while P^>=' ' do begin
    HasLT := false;
    HasGT := false;
    while P^=' ' do begin
      WR.Add(' ');
      inc(P);
    end;
    if (PWord(P)^=ord('(')+ord('*')shl 8) and
       ((@KeyWords=@PASCALKEYWORDS) or (@KeyWords=@MODULA2KEYWORDS)) then begin
      B := P;
      inc(P,2);
      while (PWord(P)^<>ord('*')+ord(')')) and (P^>=' ') do begin
        if P^='<' then HasLT := True else
        if P^='>' then HasGT := True;
        inc(P);
      end;
      if P^='*' then inc(P,2);
com:  SetToken;
      WR.Add(HTML_TAGS[false,hNavyItalic]).Add(token).Add(HTML_TAGS[true,hNavyItalic]);
      continue;
    end else
    if (P^='{') and (@KeyWords=@PASCALKEYWORDS) then begin
      B := P;
      repeat
        if P^='<' then HasLT := True else
        if P^='>' then HasGT := True;
        inc(P)
      until (P^='}') or (P^<' ');
      if P^='}' then inc(P);
      goto com;
    end; 
    B := P;
    if PWord(P)^=ord('/')+ord('/')shl 8 then begin
      while P^>=' ' do begin
        if P^='<' then HasLT := True else
        if P^='>' then HasGT := True;
        inc(P);
      end;
      goto com;
    end;
    repeat
      if CString and (P^='"') then begin
        if (P[1]='"') and (B=P) then begin
          WR.Add('""');
          inc(P,2);
          B := P;
        end else
          break;
      end else
      if P^='''' then begin
        if PasString and (P[1]='''') then begin
          WR.Add('''''');
          inc(P,2);
          B := P;
        end else
          break;
      end else
      if not (P^ in ['0'..'9','a'..'z','A'..'Z','_']) then
        break else
        inc(P);
    until P^<' ';
    SetString(token,B,P-B);
    HasLT := false;
    HasGT := false;
    B := P;
    if token<>'' then
      if (CString or PasString) and IsNumber(pointer(token)) then
        goto str2 else
      if ((@KeyWords=@MODULA2KEYWORDS) and
            IsKeyWord(KeyWords,token)) or // .MOD keys are always uppercase
         ((@KeyWords<>@MODULA2KEYWORDS) and IsKeyWord(KeyWords,UpperCase(token))) then
        WR.Add(HTML_TAGS[false,hBold]).Add(token).Add(HTML_TAGS[true,hBold]) else
        WR.Add(token);
    if CString and (P^='"') and (P[1]<>'"') then begin
      repeat
        if P^='<' then HasLT := True else
        if P^='>' then HasGT := True;
        inc(P);
      until (P^='"') or (P^<' ');
      if P^='"' then inc(P);
str:  SetToken;
str2: WR.Add(HTML_TAGS[false,hNavy]).Add(token).Add(HTML_TAGS[true,hNavy]);
    end else
    if PasString and (P^='''') and (P[1]<>'''') then begin
      repeat
        if P^='<' then HasLT := True else
        if P^='>' then HasGT := True;
        inc(P);
      until (P^='''') or (P^<' ');
      if P^='''' then inc(P);
      goto str;
    end else
    if token='' then begin
      if P^='<' then
        WR.Add(HTML_TAGS[false,hLT]) else
      if P^='>' then
        WR.Add(HTML_TAGS[false,hGT]) else
        WR.Add(P^);
      inc(P);
    end;
  end;
  if HighLight then
    WR.Add(HTML_TAGS[true,hHighlight]);
end;

procedure THTML.RtfPage;
begin
  WR.AddCRLF;
end;

function THTML.RtfPageRefToString(aBookName: string; withLink: boolean;
  BookMarkAlreadyComputed: boolean; Sequence: integer): string;
var caption: string;
begin
  result := '';
  if (aBookName='') or not withLink then
    exit;
  if not BookMarkAlreadyComputed then
    aBookName := RtfBookMarkName(aBookName);
  if Sequence<>0 then
    caption := IntToStr(abs(Sequence)) else
    caption := '...';
  result := Format(HTML_TAGS[false,hAHref],['#'+aBookName])+
    '<span class="label';
  if Sequence<0 then
    result := result+' label-primary';
  result := result+'">'+caption+'</span>'+HTML_TAGS[true,hAHref];
  if not fStringPlain then
    result := #1+result+#1;
end;

procedure THTML.RtfTitle(Title: string; LevelOffset: integer;
  withNumbers: boolean; Bookmark: string);
var i,lev: Integer;
begin
  inherited; // set TitleLevel*[] and fInRtfTitle/fBookmarkInRtfTitle
  lev := TitleLevelCurrent;
  if lev>4 then
    lev := 4;
  WR.AddCRLF.Add(HTML_TAGS[false,hH],[lev]);
  if fBookmarkInRtfTitle<>'' then
    WR.Add(RtfBookMarkString('',fBookmarkInRtfTitle,true));
  i := Pos('\line',fInRtfTitle);
  if i>0 then begin
    fInRtfTitle[i] := #0;
    WriteAsHtml(pointer(fInRtfTitle),nil);
    WR.AddShort('<br><small>');
    WriteAsHtml(@fInRtfTitle[i+5],nil);
    WR.AddShort('</small>');
  end else
    WriteAsHtml(pointer(fInRtfTitle),nil);
  WR.Add(HTML_TAGS[true,hH],[lev]).AddCRLF;
  RtfText;
end;

procedure THTML.RtfSubTitle(const Title: string);
var lev: integer;
begin
  lev := TitleLevelCurrent;
  if lev>6 then
    lev := 6;
  WR.AddCRLF.Add(HTML_TAGS[false,hH],[lev]);
  WriteAsHtml(pointer(Title),nil);
  WR.Add(HTML_TAGS[true,hH],[lev]).AddCRLF;
end;

function UTF8Encode(const winansi: string): utf8string;
const VOIDP1=ord('<')+ord('p')shl 8+ord('>')shl 16+ord('<')shl 24;
      VOIDP2=ord('/')+ord('p')shl 8+ord('>')shl 16+13 shl 24;
      VOIDP3=ord('<')+ord('p')shl 8+ord('>')shl 16+ord(' ')shl 24;
      VOIDP4=ord('<')+ord('/')shl 8+ord('p')shl 16+ord('>') shl 24;
      CRLF=$0a0d0a0d;
var len: Integer;
    S,E,P: PAnsiChar;
begin
  len := length(winansi);
  S := pointer(winansi);
  E := S+len;
  while S<E do
    if PInteger(S)^=CRLF then begin
      inc(S,2);
      dec(len,2);
    end else
    if (PInteger(S)^=VOIDP1) and (PInteger(S+4)^=VOIDP2) then begin
      inc(S,9);
      dec(len,9);
      while PWord(S)^=$a0d do begin
        inc(S,2);
        dec(len,2);
      end;
    end else
    if (PInteger(S)^=VOIDP3) and (PInteger(S+4)^=VOIDP4) then begin
      inc(S,8);
      dec(len,8);
      while PWord(S)^=$a0d do begin
        inc(S,2);
        dec(len,2);
      end;
    end else begin
      if S^>#$7f then
        inc(len);
      inc(S);
    end;
  SetLength(result,len);
  S := pointer(winansi);
  P := pointer(result);
  while S<E do
    if PInteger(S)^=CRLF then
      inc(S,2) else
    if (PInteger(S)^=VOIDP1) and (PInteger(S+4)^=VOIDP2) then begin
      inc(S,9);
      while PWord(S)^=$a0d do 
        inc(S,2);
    end else
    if (PInteger(S)^=VOIDP3) and (PInteger(S+4)^=VOIDP4) then begin
      inc(S,8);
      while PWord(S)^=$a0d do 
        inc(S,2);
    end else begin
      if S^>#$7f then begin
        P[0] := AnsiChar($C0 or (ord(S^) shr 6));
        P[1] := AnsiChar($80 or (ord(S^) and $3F));
        inc(P,2);
      end else begin
        if S^=#0 then // fix incorrect input
          P^ := ' ' else
          P^ := S^;
        inc(P);
      end;
      inc(S);
    end;
  assert(P-pointer(result)=len);
end;

procedure THTML.SaveToFile(Format: TSaveFormat; OldWordOpen: boolean);
var html: AnsiString;
begin
  RtfText;
  WR.Add(CONTENT_FOOTER);
  if Format<>fHtml then
    exit;
  if OldWordOpen then
    html := '../';
  html := UTF8Encode(SysUtils.Format(CONTENT_HEADER,[AnsiQuotedStr(fContent,'"'),
    AnsiQuotedStr(fAuthor,'"'),HtmlEncode(fTitle),html])+WR.Data);
  StringToFile(FileName,html);
end;

procedure THTML.SetLandscape(const Value: boolean);
begin
end;

function THTML.ContentAsHtml(const text: string): string;
var Temp: TStringWriter;
    lev: integer;
    saved,savedStack: THtmlTags;
begin
  Temp.Init;
  lev := Level;
  saved := Current;
  savedStack := Stack[lev];
  WriteAsHtml(pointer(text),@Temp);
  Level := lev;
  Stack[Level] := savedStack; // emulates }
  SetCurrent(@Temp);
  Current := saved;
  result := Temp.GetData;
end;

procedure THTML.SetCurrent(W: PStringWriter);
var Old, New: THtmlTags;
    Tag: THtmlTag;
begin
  New := Stack[Level];
  Old := Current;
  if New=Old then
    exit;
  for Tag := low(Tag) to high(Tag) do
    if (Tag in Old) and not (Tag in New) then
      W^.Add(HTML_TAGS[true,Tag]) else
    if not (Tag in Old) and (Tag in New) then
      W^.Add(HTML_TAGS[false,Tag]);
  Current := New;
end;

procedure THTML.WriteAsHtml(P: PAnsiChar; W: PStringWriter);
var B: PAnsiChar;
    L: integer;               
    token: AnsiString;
begin
  if P=nil then
    exit;
  if W=nil then
    W := @WR;
  while P^<>#0 do begin
    case P^ of
      #1: begin
        SetCurrent(W);
        repeat // #1...#1 is written directly with no conversion
          inc(P);
          case P^ of
          #0: exit;
          #1: break;
          end;
          W^.Add(P^);
        until false;          
      end;
      #10,#13: ; // just ignore control chars
      '{': begin
        B := P;
        repeat inc(B) until B^ in [#0,'}'];
        if B^<>#0 then begin
          if Level<high(Stack) then begin
            inc(Level);
            Stack[Level] := Stack[Level-1];
          end;
        end;
      end;
      '}': if Level>0 then dec(Level);
      '&': begin
        SetCurrent(W);
        W^.Add(HTML_TAGS[false,hAMP]);
      end;
      '<': begin
        SetCurrent(W);
        W^.Add(HTML_TAGS[false,hLT]);
      end;
      '>': begin
        SetCurrent(W);
        W^.Add(HTML_TAGS[false,hGT]);
      end;
      '\': begin
        B := P;
        repeat inc(B) until B^ in RTFEndToken;
        L := B-P-1;
        if L<=7 then begin
          if L>0 then begin
            SetString(Token,p+1,L);
            if token='b' then
              include(Stack[Level],hBold) else
            if token='b0' then
              exclude(Stack[Level],hBold) else
            if token='i' then
              include(Stack[Level],hItalic) else
            if token='i0' then
              exclude(Stack[Level],hItalic) else
            if token='ul' then
              include(Stack[Level],hUnderline) else
            if token='ul0' then
              exclude(Stack[Level],hUnderline) else
            if token='strike' then
              include(Stack[Level],hItalic) else
            if token='f1' then
              include(Stack[Level],hCode) else
            if token='f0' then
              exclude(Stack[Level],hCode) else
            if token='line' then
              W^.Add(HTML_TAGS[false,hBR]) else
            if token='pard' then
              Stack[Level] := [] else
            if token='par' then begin // should not occur normaly
              SetCurrent(W);
              W^.Add(HTML_TAGS[false,hP]);
            end else
            if token='tab' then
              W^.Add('&nbsp;') else
            if (token<>'qc') and (token<>'ql') and (token<>'qr') and
               (token<>'qj') and (token<>'shad') and 
               (PWord(token)^<>ord('c')+ord('f')shl 8) and
               (PWord(token)^<>ord('f')+ord('i')shl 8) and
               (PWord(token)^<>ord('l')+ord('i')shl 8) and
               (PWord(token)^<>ord('f')+ord('s')shl 8) and
               (PWord(token)^<>ord('s')+ord('a')shl 8) and
               (PWord(token)^<>ord('s')+ord('b')shl 8) then begin
              // ignore \qc\ql\qr\qj\cf#\fi#\li#\fs#\sa#\sb#
              OnError('Unknown token "%s"',[token]);
              W^.Add('???').Add(Token);
            end;
          end else
            if P[1]='\' then begin
              W^.Add('\');
              inc(P,2);
              continue;
            end;
          inc(P,L+1);
          if P^=#0 then
            break;
          continue;
        end;
      end;
      else begin
        if Stack[Level]<>Current then
          SetCurrent(W);
        W^.Add(P^);
      end;
    end;
    inc(P);
  end;
end;



initialization

finalization
{$ifdef DIRECTEXPORTTOWORD}
  if CoInitialized then
    CoUnInitialize;
{$endif}
end.


