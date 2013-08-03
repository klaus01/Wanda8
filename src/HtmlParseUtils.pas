{ 解析网页的公共方法，类 }
unit HtmlParseUtils;

interface

uses
  mshtml, Windows, Classes, SysUtils, StrUtils, ActiveX, Contnrs, ExtCtrls,
  SHDocVW, Graphics, Messages;

const
  WM_PROCESS_YESNODLG = WM_USER + $9801;
  WM_REMOVE_OBJ       = WM_USER + $9802;

type

  EHtmlParseException = class(Exception)
  end;

  { 表格数据支持．提供者可以是网页，excel文件或文本文件 }
  IGridDataProvider = interface
    { todo: a column name x-reference table required }
    function GetCell(AColName: string): string;
    function GetCellDef(AColName: string; const ADef: string = ''): string;
    function GetCellHtmlDef(AColName: string; const ADef: string = ''): string;
    property Cells[AColName: string]: string read GetCell; default;
  end;

  { 处理横表，整个表表示一条记录．它的奇数列指出名字，偶数列为值． }
  THorizTableParser = class(TObject, IGridDataProvider)
  private
    FCells: TStringList; // colnames & colvalues mixed
  public
    constructor Create(ATable: IHTMLTable); reintroduce;
    destructor Destroy; override;
    { 特殊表格调用．设置第一个可识别的列名．在此列之前的单元格均被裁剪掉． }
    procedure SetFirstColumn(const ACol: string);
    function GetCell(AColName: string): string;
    property Cells[AColName: string]: string read GetCell; default;
    function GetCellDef(AColName: string; const ADef: string = ''): string;
    function GetCellHtmlDef(AColName: string; const ADef: string = ''): string;
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
  end;

  { 处理交叉表，第一列为列头，第一行为表头，一对表头和列头定位一个单元格 }
  TCrossTableParser = class(TObject)
  private
    FColHeader, FRowHeader: TStringList;
    FTable: IHTMLTable;
  public
    constructor Create(ATable: IHTMLTable); reintroduce;
    destructor Destroy; override;
    function ParseHeader(ColHeader: array of string): Boolean;
    function GetCellDef(AColName, ARowName: string; const ADef: string = ''): string;
  end;

  { 处理第一列为表头(或用ParseHeaderWithCols指出表头包含的列)，以下列为表内容的表格 }
  TSimpleTableParser = class(TObject, IGridDataProvider)
  private
    FCols: TStringList; // colname, object=integer of col index
    FRowIndex: Integer;
    FTable: IHTMLTable;
    function GetCell(AColName: string): string;
    function GetRowCount: Integer;
    procedure ParseHeader;
  public
    constructor Create(ATable: IHTMLTable); reintroduce;
    destructor Destroy; override;
    { 特殊表格调用。处理一个包含给定表头的行 }
    procedure ParseHeaderWithCols(Cols: array of string);
    property Cells[AColName: string]: string read GetCell; default;
    function GetCellDef(AColName: string; const ADef: string = ''): string; overload;
    function GetCellDef(AColIndex: Integer; const ADef: string = ''): string; overload;
    function GetCellHtmlDef(AColName: string; const ADef: string = ''): string;
    function GetTableCell(AColName: string): IHTMLElement;
    function GetRow: IHTMLTableRow;
    function BackRow: Boolean;
    function NextRow: Boolean;
    procedure Reset;
    property RowCount: Integer read GetRowCount;
    function EOT: Boolean;
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
  end;

  { 处理导出文件的基类．纯虚类．接口尽量简单． }
  TCommonExchangeFileParserBase = class(TObject, IGridDataProvider)
  public
    constructor Create(AFileName: string); virtual; abstract;
    destructor Destroy; override; abstract;
    function TextExists(AText: string): Boolean; virtual; abstract;
    function ParseHeader(ARequiredCols: array of string): Boolean; virtual; abstract;
    function NextRow: Boolean; virtual; abstract;
    function EOT: Boolean; virtual; abstract;
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
    function GetCell(AColName: string): string; virtual; abstract;
    function GetCellDef(AColName: string; const ADef: string = ''): string; virtual; abstract;
    function GetCellHtmlDef(AColName: string; const ADef: string = ''): string; virtual; abstract;
    property Cells[AColName: string]: string read GetCell; default;
  end;

  { 处理逗号分隔的CSV文件 }
  TCSVGridParser = class(TCommonExchangeFileParserBase)
  private
    FCols: TStringList;
    FLines: TStringList;
    FLineSplitted: TStringList;
    FCurLine: Integer;
    FLastLine: Integer;
  protected
    procedure LineSplittedRequired;
    function RefineCell(S: string): string;
  public
    constructor Create(AFileName: string); reintroduce;
    destructor Destroy; override;
    function TextExists(AText: string): Boolean; override;
    function ParseHeader(ARequiredCols: array of string): Boolean; override;
    function NextRow: Boolean; override;
    function EOT: Boolean; override;
    function GetCell(AColName: string): string; override;
    function GetCellDef(AColName: string; const ADef: string = ''): string; override;
    property Cells; default;
  end;

  TMultiPageParseHelper = class;
  TOnGetPageOKProc = procedure (ASender: TMultiPageParseHelper; var PageOK: Boolean) of object;
  TOnAnalysePageProc = procedure (ASender: TMultiPageParseHelper) of object;
  TOnGetIsLastPageProc = procedure (ASender: TMultiPageParseHelper; var IsLastPage: Boolean) of object;
  TOnNextPageProc = procedure (ASender: TMultiPageParseHelper) of object;
  TOnAllPageOK = procedure (ASender: TMultiPageParseHelper) of object;
  TOnGetHtmlProc = procedure (ASender: TMultiPageParseHelper; var aNewHtml: string) of object;

  { 处理分多页的表格/页面 }
  TMultiPageParseHelper = class
  private
    FTimer: TTimer;
    FLastHtml: string;
    FDocument: IHTMLDocument2;
    FPageIndex: Integer;
    FOnAllPageOK: TOnAllPageOK;
    FOnAnalysePage: TOnAnalysePageProc;
    FOnGetIsLastPage: TOnGetIsLastPageProc;
    FOnGetPageOK: TOnGetPageOKProc;
    FOnNextPage: TOnNextPageProc;
    FIfCompareHtml: Boolean;
    FOnGetHtml: TOnGetHtmlProc;
  protected
    function IsPageOK: Boolean; virtual;
    procedure AnalysePage; virtual;
    function IsLastPage: Boolean; virtual;
    procedure NextPage; virtual;
    procedure AllPageOK; virtual;
    function GetHtml: string; virtual;
    procedure OnTimer(Sender: TObject);
  public
    constructor Create; reintroduce;
    destructor Destroy; override;
    { 当前分析的页 }
    property PageIndex: Integer read FPageIndex;
    { 开始分析 }
    procedure Start(ADoc: IHTMLDocument2);
    { 停止分析 }
    procedure Stop;
    { 当前页是否加载完成 }
    property OnGetPageOK: TOnGetPageOKProc read FOnGetPageOK write FOnGetPageOK;
    { 分析当前页 }
    property OnAnalysePage: TOnAnalysePageProc read FOnAnalysePage write FOnAnalysePage;
    { 是否是最后一页 }
    property OnGetIsLastPage: TOnGetIsLastPageProc read FOnGetIsLastPage write FOnGetIsLastPage;
    { 前往下一页 }
    property OnNextPage: TOnNextPageProc read FOnNextPage write FOnNextPage;
    { 所有页分析完成 }
    property OnAllPageOK: TOnAllPageOK read FOnAllPageOK write FOnAllPageOK;
    { 是否比较两次页面 (相同则不进行页面分析) }
    property IfCompareHtml: Boolean read FIfCompareHtml write FIfCompareHtml;
    { 比较网页时，可由外部确定比较网页的内容（当网页的变化是在Frame中时，就得用这个） }
    property OnGetHtml: TOnGetHtmlProc read FOnGetHtml write FOnGetHtml;
  end;

  { 全局类,内部使用.自动捕捉IE弹出的脚本错误对话框,并且对"是否继续运行脚本"点它的"是"按钮
    (只对主线程有效) }
  TIEPopupDialogWatcher = class
  private
    FHandle: HWND;
  protected
    procedure WndProc(var Msg: TMessage);
    procedure OnProcessYesNoDlg(var Msg: TMessage); message WM_PROCESS_YESNODLG;
  public
    constructor Create;
    destructor Destroy; override;
    property Handle: HWND read FHandle;
  end;

  TNavigateAndCallFunc = procedure(AWB: TWebBrowser; var Done: Boolean) of object;

// 得到某一帧的document
function GetFrameEleColl(FrameNum: OleVariant; HtmlDoc: IHTMLDocument2):
   IHTMLdocument2;

// 比较两个字符串是否互相包含
function CompareInnerText(s1,s2:string;ACase:boolean):boolean;

// 触发innertext=xxx的链接
function HrefClick(SoucEleColl:OleVariant;Const AText:string;ACase:Boolean=false):boolean;

// 触发href=xxx的链接
function HrefClickByHrefStr(SoucEleColl:OleVariant;Const AText:string;ACase:Boolean=false):boolean;

// 根据某一属性查找对象
function LocateObject(SoucEleColl:OleVariant;Const TgName,AText,AttrStr:string):OleVariant;

// 从一个窗口中查找它容纳的IHTMLDocument2对象（会自动查找子窗口）  IE5.5
function GetIHTMLDocumentFromHWND(const H: HWND): IHTMLDocument2;

{ 查找给出id的frame(递归) }
function FindFrameByID(Doc: IHTMLDocument2; const FrameID: string): IHTMLWindow2;

{ 获取给出id的select项的选中文字 }
function GetSelectOptionText(ADoc: IHTMLDocument2; const SelectID: string): string;

{ 查找包含所有给出列名的table(注：既查找纵表，也查找横表) }
function FindTableWithColumns(ADoc: IHTMLDocument2; Cols: array of string; const aHeadTag: string = 'td'; const aTabIndex: Integer = 0): IHTMLTable;
{ 处理带','号的字符串币种 }
function MWStrToCurrDef(S: string; ADef: Currency): Currency;
{ 获得某select框的所有选项和值列表．ATexts,AValues由调用方创建 }
function GetSelectOptions(ADoc: IHTMLDocument2; ASelectId: string; ATexts, AValues: TStrings): Boolean;
{ 从IHTMLIMGElement中抓取图像.Bmp由调用方创建和释放. }
function CaptureHtmlImg(const AImg: IHTMLIMGElement; var Bmp: TBitmap): Boolean;
{ 设置select框的选项 }
function SetSelectOption(ADoc: IHTMLDocument2; ASelectId: string; AOption: string): Boolean;
function SetSelectOptionByValue(ADoc: IHTMLDocument2; ASelectId: string; AValue: string): Boolean;
{ 设置input框的文本 }
function SetInputTagText(ADoc: IHTMLDocument2; AInputId: string; AText: string): Boolean;

{ 浏览器访问网页，在其完成后调用给出方法．若方法返回未完成，则持续调用直到完成；若超时无结果，则调用超时方法．
  * !!! * 在ACallback或ATimeoutFunc对象析构时必须调用UnregisterCallerObject方法移除所有相关回调对象。 }
{ NOTE: This function has been prooved very *BUGGY*, use CallUntilDone instead. }
procedure NavigateAndCall(AWB: TWebBrowser; AUrl: string; ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent = nil); deprecated;
{ 在定时器中不断尝试调用ACallback，直到ACallback返回Done=TRUE或者超时。超时则调用ATimeoutFunc，
  * !!! * 在ACallback或ATimeoutFunc对象析构时必须调用UnregisterCallerObject方法移除所有相关回调对象。 }
procedure CallUntilDone(AWB: TWebBrowser; ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent = nil);
{ 移除对象所有相关回调对象 }
procedure UnregisterCallerObject(AObject: TObject);


{ 检查表格是否包含给出的所有单元格 }
function IsTableContainCells(const ATable: IHTMLTable; const ACells: array of string): Boolean;
{ 查找给出的tag和class的元素．找不到返回nil }
function FindElemByTagAndClass(ADoc: IHTMLDocument2; const ATagName, AClassName: string): IHTMLElement;
function FindElemByTagAndID(ADoc: IHTMLDocument2; const ATagName, AID: string): IHTMLElement;
function FindElemByTagAndText(ADoc: IHTMLDocument2; const ATagName, AText: string): IHTMLElement;
function FindElemByTagAndTitle(ADoc: IHTMLDocument2; const ATagName, ATitle: string): IHTMLElement;
{ 根据value查找Input }
function FindButtonByValue(ADoc: IHTMLDocument2; const aValue: string): IHTMLButtonElement;
function FindInputByValue(ADoc: IHTMLDocument2; const aValue: string): IHTMLInputElement;
function FindInputByName(ADoc: IHTMLDocument2; const aName: string): IHTMLInputElement;
{ 根据src属性查找img对象 }
function FindIMGBySRC(ADoc: IHTMLDocument2; const aSRC: string): IHTMLImgElement;
{ 计算HTML元素的区域。相对于Document的区域 }
function CalcElementClientRect(AElem: IHTMLElement): TRect;
{ 计算HTML元素的区域，方法2 }
function CalcElementClientRect2(AElem, ABody: IHTMLElement): TRect;
{ 调用一个IDispatch对象的方法 }
function SafeFireEvent(ADisp: IDispatch; AEventName: string; Params: array of const): Boolean;
{ 使用post方法提交数据给iframe }
procedure FramePostData(AWB: TWebBrowser; AUrl: string; APostData: string; AFrameId: string);
{ 设置自定义控件文本。它必须实现IOleWindow接口，并且是通过GetFocus来判断是否被激活的。 }
procedure SetCustomActiveXControlText(AElem: IHTMLElement; AText: string; ADllName: string);

{ 调用ImmDisableIME API禁止本线程使用IME。插件创建Browser窗口前调用。 }
procedure DisableCurrentThreadIME;

implementation

uses
  CommCtrl, Forms, Variants, IMM, SyncObjs;

type
  TNavigateAndCallObj = class
  private
    FWB: TWebBrowser;
    FUrl: string;
    FTimer: TTimer;
    FTimerTimeout: TTimer;
    FOldDocumentComplete: TWebBrowserDocumentComplete;
    FCallback: TNavigateAndCallFunc;
    FOnTimeout: TNotifyEvent;
  private
    procedure OnTimer(Sender: TObject);
    procedure OnTimeout(Sender: TObject);
  public
    constructor Create(AWB: TWebBrowser; AUrl: string; ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent = nil); reintroduce;
    destructor Destroy; override;
    procedure OnDocumentComplete(Sender: TObject; const pDisp: IDispatch; var URL: OleVariant);
  end;

  TCallUntilDoneObj = class
  private
    FWB: TWebBrowser;
    FTimer: TTimer;
    FCallback: TNavigateAndCallFunc;
    FTimeCount: Integer;
    FTimeoutFunc: TNotifyEvent;
  private
    procedure OnTimer(Sender: TObject);
  public
    constructor Create(AWB: TWebBrowser; ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent);
    destructor Destroy; override;
  end;

  TTempObjList = class(TList)
  private
    FHandle: HWND;
    procedure WndProc(var Msg: TMessage);
  public
    constructor Create;
    { **Note** AObj is type of TCallUntilDoneObj or TNavigateAndCallObj }
    procedure Remove(AObj: TObject); reintroduce;
    { * Remove all TCallUntilDoneObj/TNavigateAndCallObj which contain object pointers to the caller object * }
    procedure OnCallerDestroyed(ACaller: TObject);
    destructor Destroy; override;
    procedure Add(AObj: TObject); reintroduce;
    property Handle: HWND read FHandle;
  end;

var
  l_IEPopupDialogWatcher: TIEPopupDialogWatcher;
  l_OldCallWndProcHook: HHOOK = 0;
  l_TempObjList: TTempObjList;

//得到某一帧的document
function GetFrameEleColl(FrameNum: OleVariant; HtmlDoc: IHTMLDocument2):
   IHTMLdocument2;
var
 Nw:IHTMLWindow2;
 HF:IHTMLFramesCollection2;
 spDisp : IDispatch;
begin
 HF:=HtmlDoc.frames;
 spDisp:=HF.item(FrameNum);
 if SUCCEEDED(spDisp.QueryInterface(IHTMLWindow2 ,Nw))then
 begin
   result:=(Nw.document as IHTMLdocument2);
 end;
end;

function CompareInnerText(s1,s2:string;ACase:boolean):boolean;
begin
  result:=false;
  if ACase then
   begin
    if pos(s2,s1)>0 then
     result:=true;
   end
  else
    if (Trim(s1)=Trim(s2)) then
     result:=true;
end;

//触发innertext=xxx的链接
function HrefClick(SoucEleColl:OleVariant;Const AText:string;ACase:Boolean=false):boolean;
var
 Item,FAll:OleVariant;
 i:integer;
begin
     result:=false;
     FAll:=SoucEleColl.Tags('A');
     for i:=0 to FAll.length-1 do
       begin
        item:=FAll.item(i);
        if CompareInnerText(item.innerText,AText,ACase) then
         begin
           item.click;
           result:=true;
           break;
         end;
       end

end;

//触发href=xxx的链接
function HrefClickByHrefStr(SoucEleColl:OleVariant;Const AText:string;ACase:Boolean=false):boolean;
var
 Item,FAll:OleVariant;
 i:integer;
begin
     result:=false;
     FAll:=SoucEleColl.Tags('A');
     for i:=0 to FAll.length-1 do
       begin
        item:=FAll.item(i);
        if CompareInnerText(item.href,AText,ACase) then
         begin
           item.click;
           result:=true;
           break;
         end;
       end

end;

//根据某一属性查找对象
function LocateObject(SoucEleColl:OleVariant;Const TgName,AText,AttrStr:string):OleVariant;
var
 Item,FAll:OleVariant;
 i:integer;
begin
   FAll:=SoucEleColl.Tags(TgName);
    for i:=0 to FAll.length-1 do
       begin
        item:=FAll.item(i);
        if Trim(item.getAttribute(AttrStr))=Trim(AText) then
         begin
           result:=item;
           break;
         end;
       end;
end;

type
  PHWND = ^HWND;

function __EnumChildWindowsProc(H: HWND; lp: LPARAM): LongBool; stdcall;
var
  ClsName: string;
begin
  SetLength(ClsName, 32);
  SetLength(ClsName, GetClassName(H, PChar(ClsName), 32));
  OutputDebugString(PChar(ClsName));
  if ClsName = 'Internet Explorer_Server' then
  begin
    PHWND(lp)^ := H;
    Result := False;
    Exit;
  end;
  Result := True;
end;

function GetIHTMLDocumentFromHWND(const H: HWND): IHTMLDocument2;
type
  TObjectFromLresultFunc = function (lr: LRESULT; const riid: TIID; wp: WPARAM; var pObj: IUnknown): HRESULT; stdcall;
var
  LibHandle: HMODULE;
  ObjectFromLresult: TObjectFromLresultFunc;
  lr: Cardinal;
  Msg: UINT;
  Obj: IUnknown;
  TargetWnd: HWND;
  ClsName: string;
begin
  Result := nil;
  TargetWnd := 0;
  SetLength(ClsName, 32);
  SetLength(ClsName, GetClassName(H, PChar(ClsName), 32));
  if ClsName = 'Internet Explorer_Server' then TargetWnd := H
  else EnumChildWindows(H, @__EnumChildWindowsProc, Integer(@TargetWnd));
  if TargetWnd = 0 then Exit;
  CoInitialize(nil);
  try
    LibHandle := LoadLibrary('oleacc.dll');
    if LibHandle = 0 then Exit;
    try
      Msg := RegisterWindowMessage('WM_HTML_GETOBJECT');
      SendMessageTimeout(TargetWnd, Msg, 0, 0, SMTO_ABORTIFHUNG, 1000, lr);
      @ObjectFromLresult := GetProcAddress(LibHandle, 'ObjectFromLresult');
      if @ObjectFromLresult = nil then Exit;
      ObjectFromLresult(lr, IID_IHTMLDocument2, 0, Obj);
      if (Obj <> nil) and Supports(Obj, IID_IHTMLDocument2) then Result := Obj as IHTMLDocument2;
    finally
      FreeLibrary(LibHandle);
    end;
  finally
    CoUninitialize;
  end;
end;

{ 查找给出id的frame(递归) }
function FindFrameByID(Doc: IHTMLDocument2; const FrameID: string): IHTMLWindow2;
var
  I: Integer;
  ItemID: OleVariant;
  Elem: IHTMLWindow2;
begin
  Result := nil;
  for I := 0 to Doc.frames.length - 1 do
  begin
    ItemID := I;
    Elem := IDispatch(Doc.frames.item(ItemID)) as IHTMLWindow2;
    if SameText(Elem.name, FrameID) then
    begin
      Result := Elem;
      Exit;
    end;
    Result := FindFrameByID(Elem.document, FrameID);
    if Result <> nil then Exit;
  end;
end;

{ 获取给出id的select项的选中文字 }
function GetSelectOptionText(ADoc: IHTMLDocument2; const SelectID: string): string;
var
  Elem: IHTMLElement;
  Select: IHTMLSelectElement;
begin
  Elem := (ADoc.all as IHTMLElementCollection).item(SelectID, 0) as IHTMLElement;
  if not Supports(Elem, IHTMLSelectElement) then Exit;
  Select := Elem as IHTMLSelectElement;
  if Select.selectedIndex < 0 then Exit;
  try
    Elem := Select.item(Select.selectedIndex, 0) as IHTMLElement;
  except
    Exit;
  end;
  if not Supports(Elem, IHTMLOptionElement) then Exit;
  Result := (Elem as IHTMLOptionElement).text;
end;

function FindTableWithColumns(ADoc: IHTMLDocument2; Cols: array of string; const aHeadTag: string; const aTabIndex: Integer): IHTMLTable;

  function FindSubTables(AAll: IHTMLElementCollection; Cols: array of string; const aTabIndex: Integer): IHTMLTable;
  var
    Tags, Cells: IHTMLElementCollection;
    Table: IHTMLTable;
    Elem: IHTMLElement;
    I, J, Index, N: Integer;
    SL: TStringList;
    vTabIndex: Integer;
  begin
    Result := nil;
    vTabIndex := aTabIndex;
    Tags := AAll.tags('table') as IHTMLElementCollection;
    SL := TStringList.Create;
    try
      SL.Sorted := True;
      for I := Low(Cols) to High(Cols) do SL.Add(Cols[I]);
      for I := 0 to Tags.length - 1 do
      begin
        Table := Tags.item(I, 0) as IHTMLTable;
        for J := 0 to SL.Count - 1 do SL.Objects[J] := TObject(0);
        N := SL.Count;
        Cells := ((Table as IHTMLElement).all as IHTMLElementCollection).tags(aHeadTag) as IHTMLElementCollection;
        for J := 0 to Cells.length - 1 do
        begin
          Elem := Cells.item(J, 0) as IHTMLElement;
          Index := SL.IndexOf(Trim(Elem.innerText));
          if Index < 0 then Continue;
          if Integer(SL.Objects[Index]) <> 0 then Continue;
          SL.Objects[Index] := TObject(1);
          Dec(N);
          if N = 0 then
          begin
            if vTabIndex > 0 then
            begin
              Dec(vTabIndex);
              Break;
            end;
            Result := Table;
            // 找尽可能小的table
            Table := FindSubTables((Table as IHTMLElement).all as IHTMLElementCollection, Cols, aTabIndex);
            if Table <> nil then Result := Table;
            Exit;
          end;
        end;
      end;
    finally
      FreeAndNil(SL);
    end;
  end;
begin
  Result := FindSubTables(ADoc.all, Cols, aTabIndex);
end;

{ TSimpleTableParser }

function TSimpleTableParser.BackRow: Boolean;
var
  Row: IHTMLTableRow;
begin
  Result := False;
  if FRowIndex <= 0 then Exit;
  Dec(FRowIndex);
  if FRowIndex <= 0 then Exit;
  Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
  if Row.cells.length <> FCols.Count then Result := BackRow // not a valid row
  else Result := True;
end;

constructor TSimpleTableParser.Create(ATable: IHTMLTable);
begin
  FRowIndex := 0;
  FCols := TStringList.Create;
  FTable := ATable;
  ParseHeader;
end;

destructor TSimpleTableParser.Destroy;
begin
  FreeAndNil(FCols);
  inherited;
end;

function TSimpleTableParser.EOT: Boolean;
begin
  Result := FRowIndex >= FTable.rows.length;
end;

function TSimpleTableParser.GetCell(AColName: string): string;
var
  Row: IHTMLTableRow;
  Index: Integer;
  Elem: IHTMLElement;
begin
  if FRowIndex >= FTable.rows.length then raise EHtmlParseException.Create('index out of range');
  Index := FCols.IndexOf(AColName);
  if Index < 0 then raise EHtmlParseException.CreateFmt('column %s not found', [AColName]);
  Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
  Elem := Row.cells.item(Index, 0) as IHTMLElement;
  Result := Trim(Elem.innerText);
end;

function TSimpleTableParser.GetCellDef(AColName: string;
  const ADef: string): string;
var
  Row: IHTMLTableRow;
  Index: Integer;
  Elem: IHTMLElement;
begin
  if FRowIndex >= FTable.rows.length then Result := ADef
  else
  begin
    Index := FCols.IndexOf(AColName);
    if Index < 0 then Result := ADef
    else
    begin
      Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
      Elem := Row.cells.item(Index, 0) as IHTMLElement;
      Result := Trim(Elem.innerText);
    end;
  end;
end;

function TSimpleTableParser.GetCellDef(AColIndex: Integer;
  const ADef: string): string;
var
  Row: IHTMLTableRow;
  Index: Integer;
  Elem: IHTMLElement;
begin
  if FRowIndex >= FTable.rows.length then Result := ADef
  else
  begin
    Index := AColIndex;
    if Index < 0 then Result := ADef
    else
    begin
      Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
      Elem := Row.cells.item(Index, 0) as IHTMLElement;
      Result := Trim(Elem.innerText);
    end;
  end;
end;

function TSimpleTableParser.GetCellHtmlDef(AColName: string;
  const ADef: string): string;
var
  Row: IHTMLTableRow;
  Index: Integer;
  Elem: IHTMLElement;
begin
  if FRowIndex >= FTable.rows.length then Result := ADef
  else
  begin
    Index := FCols.IndexOf(AColName);
    if Index < 0 then Result := ADef
    else
    begin
      Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
      Elem := Row.cells.item(Index, 0) as IHTMLElement;
      Result := Trim(Elem.innerHTML);
    end;
  end;
end;

function TSimpleTableParser.GetRow: IHTMLTableRow;
begin
  if FRowIndex >= FTable.rows.length then
    Result := nil
  else
    Result := FTable.rows.item(FRowIndex, FRowIndex) as IHTMLTableRow;
end;

function TSimpleTableParser.GetRowCount: Integer;
begin
  Result := FTable.rows.length - 1;
end;

function TSimpleTableParser.GetTableCell(AColName: string): IHTMLElement;
var
  Index: Integer;
  Row: IHTMLTableRow;
begin
  if FRowIndex >= FTable.rows.length then Result := nil
  else
  begin
    Index := FCols.IndexOf(AColName);
    if Index < 0 then Result := nil
    else
    begin
      Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
      Result := Row.cells.item(Index, 0) as IHTMLElement;
    end;
  end;
end;

function TSimpleTableParser.NextRow: Boolean;
var
  Row: IHTMLTableRow;
begin
  Result := False;
  if FRowIndex >= FTable.rows.length then Exit;
  Inc(FRowIndex);
  if FRowIndex >= FTable.rows.length then Exit;
  Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
  if Row.cells.length <> FCols.Count then Result := NextRow // not a valid row
  else Result := True;
end;

procedure TSimpleTableParser.ParseHeader;
var
  I: Integer;
  Row: IHTMLTableRow;
  Elem: IHTMLElement;
begin
  Row := FTable.rows.item(0, 0) as IHTMLTableRow;
  FCols.Clear;
  for I := 0 to Row.cells.length - 1 do
  begin
    Elem := Row.cells.item(I, 0) as IHTMLElement;
    FCols.Add(Trim(Elem.innerText));
  end;
  Reset;
end;

procedure TSimpleTableParser.ParseHeaderWithCols(Cols: array of string);
var
  I, J, K, L: Integer;
  Row: IHTMLTableRow;
  S: string;
  Elem: IHTMLElement;
  Found: Boolean;
begin
  FCols.Clear;
  for I := 0 to FTable.rows.length - 1 do
  begin
    Row := FTable.rows.item(I, 0) as IHTMLTableRow;
    for J := Low(Cols) to High(Cols) do
    begin
      S := Cols[J];
      Found := False;
      for K := 0 to Row.cells.length - 1 do
      begin
        Elem := Row.cells.item(K, 0) as IHTMLElement;
        if SameText(Trim(Elem.innerText), S) then
        begin
          Found := True;
          Break;
        end;
      end;
      if not Found then Break
      else if J = High(Cols) then
      begin
        // the last one found, parse then row
        for L := 0 to Row.cells.length - 1 do
        begin
          Elem := Row.cells.item(L, 0) as IHTMLElement;
          FCols.Add(Trim(Elem.innerText));
        end;
        Reset;
        Exit;
      end;
    end;
  end;
end;

function TSimpleTableParser.QueryInterface(const IID: TGUID;
  out Obj): HResult;
begin
  Result := E_NOTIMPL;
end;

procedure TSimpleTableParser.Reset;
var
  Row: IHTMLTableRow;
begin
  FRowIndex := 0;
  // 第一个满足条件的行:表头
  repeat
    if FRowIndex >= FTable.rows.length then Exit;
    Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
    if Row.cells.length = FCols.Count then Break;
    Inc(FRowIndex);
  until False;
  Inc(FRowIndex);
  // 第二个满足条件的行:第一条数据
  repeat
    if FRowIndex >= FTable.rows.length then Exit;
    Row := FTable.rows.item(FRowIndex, 0) as IHTMLTableRow;
    if Row.cells.length = FCols.Count then Break;
    Inc(FRowIndex);
  until False;
end;

function MWStrToCurrDef(S: string; ADef: Currency): Currency;
begin
  S := AnsiReplaceText(S, '$', '');
  S := AnsiReplaceText(S, '＄', '');
  S := AnsiReplaceText(S, '￥', '');
  S := AnsiReplaceText(S, '元', '');
  S := AnsiReplaceText(S, ' ', '');
  Result := StrToCurrDef(AnsiReplaceText(S, ',', ''), ADef);
end;

function TSimpleTableParser._AddRef: Integer;
begin
  Result := 0;
end;

function TSimpleTableParser._Release: Integer;
begin
  Result := 0;
end;

{ TMultiPageParseHelper }

procedure TMultiPageParseHelper.AllPageOK;
begin
  if Assigned(FOnAllPageOK) then FOnAllPageOK(Self);
end;

procedure TMultiPageParseHelper.AnalysePage;
begin
  if Assigned(FOnAnalysePage) then FOnAnalysePage(Self);
end;

constructor TMultiPageParseHelper.Create;
begin
  FDocument := nil;
  FTimer := TTimer.Create(nil);
  FTimer.Interval := 1000;
  FTimer.Enabled := False;
  FTimer.OnTimer := OnTimer;
  FIfCompareHtml := True;
  FLastHtml := '';
end;

destructor TMultiPageParseHelper.Destroy;
begin
  FreeAndNil(FTimer);
  FDocument := nil;
  inherited;
end;

function TMultiPageParseHelper.GetHtml: string;
begin
  Result := '';
  if Assigned(FOnGetHtml) then
    FOnGetHtml(Self, Result)
  else if FDocument.body <> nil then
    Result := FDocument.body.innerHTML;
end;

function TMultiPageParseHelper.IsLastPage: Boolean;
begin
  if Assigned(FOnGetIsLastPage) then FOnGetIsLastPage(Self, Result)
  else raise EHtmlParseException.Create('not assigned OnGetIsLastPage!');
end;

function TMultiPageParseHelper.IsPageOK: Boolean;
begin
  if Assigned(FOnGetPageOK) then FOnGetPageOK(Self, Result)
  else raise EHtmlParseException.Create('not assigned OnGetPageOK!');
end;

procedure TMultiPageParseHelper.NextPage;
begin
  if Assigned(FOnNextPage) then FOnNextPage(Self)
  else raise EHtmlParseException.Create('not assigned OnNextPage!');
end;

procedure TMultiPageParseHelper.OnTimer(Sender: TObject);
var
  Done: Boolean;
  NewHtml: string;
  Succ: Boolean;
begin
  Done := False;
  FTimer.Enabled := False;
  try
    if IsPageOK then
    begin
      NewHtml := GetHtml;
      if FIfCompareHtml and (FLastHtml <> '') and (FLastHtml = NewHtml) then Exit;
      AnalysePage;
      FLastHtml := NewHtml;
      if IsLastPage then
      begin
        AllPageOK;
        Done := True;
      end
      else
      begin
        Inc(FPageIndex);
        Succ := False;
        repeat
          try
            // user can raise an EAbort exception to indicate that operation failed
            NextPage;
            Succ := True;
          except
            on E: EHtmlParseException do raise
            else Succ := False;
          end;
          if not Succ then Application.ProcessMessages;
        until Succ;
      end;
    end;
  finally
    if not Done then FTimer.Enabled := True;
  end;
end;

procedure TMultiPageParseHelper.Start(ADoc: IHTMLDocument2);
begin
  FDocument := ADoc;
  FTimer.Enabled := True;
  FPageIndex := 0;
  FLastHtml := '';
end;

procedure TMultiPageParseHelper.Stop;
begin
  FTimer.Enabled := False;
end;

{ TNavigateAndCallObj }

constructor TNavigateAndCallObj.Create(AWB: TWebBrowser; AUrl: string;
  ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent);
begin
  FWB := AWB;
  FCallback := ACallback;
  FUrl := AUrl;
  FOldDocumentComplete := FWB.OnDocumentComplete;
  FWB.OnDocumentComplete := OnDocumentComplete;
  FWB.Navigate(FUrl);
  FTimer := TTimer.Create(nil);
  FTimer.Enabled := False;
  FTimer.Interval := 2000;
  FTimer.OnTimer := OnTimer;
  FTimerTimeout := TTimer.Create(nil);
  FTimerTimeout.Enabled := True;
  FTimerTimeout.Interval := 80000;
  FTimerTimeout.OnTimer := OnTimeout;
  FOnTimeout := ATimeoutFunc;
end;

destructor TNavigateAndCallObj.Destroy;
begin
  FWB.OnDocumentComplete := FOldDocumentComplete;
  FreeAndNil(FTimer);
  FreeAndNil(FTimerTimeout);
  l_TempObjList.Remove(Self);
  inherited;
end;

procedure TNavigateAndCallObj.OnDocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
var
  Done: Boolean;
begin
  if Assigned(FOldDocumentComplete) then FOldDocumentComplete(Sender, pDisp, URL);
  Done := False;
  FTimer.Enabled := False;
  FTimerTimeout.Enabled := False;
  FCallback(FWB, Done);
  if Done then PostMessage(l_TempObjList.Handle, WM_REMOVE_OBJ, 0, Integer(Self)) //Self.Free
  else
  begin
    FTimer.Enabled := True;
    FTimerTimeout.Enabled := True;
  end;
end;

procedure TNavigateAndCallObj.OnTimeout(Sender: TObject);
begin
  FTimer.Enabled := False;
  FTimerTimeout.Enabled := False;
  FWB.OnDocumentComplete := FOldDocumentComplete;
  if Assigned(FOnTimeout) then FOnTimeout(Self);
  Self.Free;
end;

procedure TNavigateAndCallObj.OnTimer(Sender: TObject);
var
  Done: Boolean;
begin
  FTimer.Enabled := False;
  FTimerTimeout.Enabled := False;
  FCallback(FWB, Done);
  if Done then Self.Free
  else
  begin
    FTimer.Enabled := True;
    FTimerTimeout.Enabled := True;
  end;
end;

procedure NavigateAndCall(AWB: TWebBrowser; AUrl: string; ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent);
begin
  l_TempObjList.Add(TNavigateAndCallObj.Create(AWB, AUrl, ACallback, ATimeoutFunc));
end;

procedure CallUntilDone(AWB: TWebBrowser; ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent);
begin
  l_TempObjList.Add(TCallUntilDoneObj.Create(AWB, ACallback, ATimeoutFunc));
end;

function FindElemByTagAndClass(ADoc: IHTMLDocument2; const ATagName, AClassName: string): IHTMLElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLElement;
begin
  Result := nil;
  Tags := ADoc.all.tags(ATagName) as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLElement;
    if SameText(Elem.className, AClassName) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindElemByTagAndID(ADoc: IHTMLDocument2; const ATagName, AID: string): IHTMLElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLElement;
begin
  Result := nil;
  Tags := ADoc.all.tags(ATagName) as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLElement;
    if SameText(Elem.id, AID) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindElemByTagAndText(ADoc: IHTMLDocument2; const ATagName, AText: string): IHTMLElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLElement;
begin
  Result := nil;
  Tags := ADoc.all.tags(ATagName) as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLElement;
    if SameText(Elem.innerText, aText) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindElemByTagAndTitle(ADoc: IHTMLDocument2; const ATagName, ATitle: string): IHTMLElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLElement;
begin
  Result := nil;
  Tags := ADoc.all.tags(ATagName) as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLElement;
    if SameText(Elem.title, ATitle) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindButtonByValue(ADoc: IHTMLDocument2; const aValue: string): IHTMLButtonElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLButtonElement;
begin
  Result := nil;
  Tags := ADoc.all.tags('button') as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLButtonElement;
    if SameText(Elem.value, aValue) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindInputByValue(ADoc: IHTMLDocument2; const aValue: string): IHTMLInputElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLInputElement;
begin
  Result := nil;
  Tags := ADoc.all.tags('input') as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLInputElement;
    if SameText(Elem.value, aValue) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindInputByName(ADoc: IHTMLDocument2; const aName: string): IHTMLInputElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLInputElement;
begin
  Result := nil;
  Tags := ADoc.all.tags('input') as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLInputElement;
    if SameText(Elem.name, aName) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function FindIMGBySRC(ADoc: IHTMLDocument2; const aSRC: string): IHTMLImgElement;
var
  Tags: IHTMLElementCollection;
  I: Integer;
  Elem: IHTMLImgElement;
begin
  Result := nil;
  Tags := ADoc.all.tags('img') as IHTMLElementCollection;
  if (Tags = nil) or (Tags.length = 0) then Exit;
  for I := 0 to Tags.length - 1 do
  begin
    Elem := Tags.item(I, 0) as IHTMLImgElement;
    if SameText(Elem.src, aSRC) then
    begin
      Result := Elem;
      Exit;
    end;
  end;
end;

function GetSelectOptions(ADoc: IHTMLDocument2; ASelectId: string; ATexts, AValues: TStrings): Boolean;
var
  Elem: IHTMLElement;
  Select: IHTMLSelectElement;
  I: Integer;
begin
  ATexts.Clear;
  AValues.Clear;
  Result := False;
  Elem := ADoc.all.item(ASelectId, 0) as IHTMLElement;
  if not Supports(Elem, IHTMLSelectElement) then Exit;
  Select := Elem as IHTMLSelectElement;
  for I := 0 to Select.length - 1 do
  begin
    Elem := Select.item(I, 0) as IHTMLElement;
    ATexts.Add(Trim(Elem.innerText));
    AValues.Add(Trim((Elem as IHTMLOptionElement).value));
  end;
  Result := True;
end;

function IsTableContainCells(const ATable: IHTMLTable; const ACells: array of string): Boolean;
var
  TDs: IHTMLElementCollection;
  Elem: IHTMLElement;
  I, J, N: Integer;
  FoundArray: array of Boolean;
begin
  Result := False;
  TDs := ((ATable as IHTMLElement).all as IHTMLElementCollection).tags('td') as IHTMLElementCollection;
  if TDs = nil then Exit;
  if Length(ACells) = 0 then Exit;
  SetLength(FoundArray, Length(ACells));
  for I := 0 to High(FoundArray) do FoundArray[I] := False;
  N := Length(ACells);
  for J := 0 to TDs.length - 1 do
  begin
    Elem := TDs.item(J, 0) as IHTMLElement;
    if Elem = nil then Continue;
    for I := 0 to High(ACells) do
    begin
      if FoundArray[I] then Continue; // already found
      if not SameText(Trim(Elem.innerText), ACells[I]) then Continue;
      FoundArray[I] := True;
      Dec(N);
      if N = 0 then // all found
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

type
// *********************************************************************//
// Interface: IHTMLElementRender
// Flags:     (0)
// GUID:      {3050F669-98B5-11CF-BB82-00AA00BDCE0B}
// *********************************************************************//
  IHTMLElementRender = interface(IUnknown)
    ['{3050F669-98B5-11CF-BB82-00AA00BDCE0B}']
    function DrawToDC(hdc: HDC): HResult; stdcall;
    function SetDocumentPrinter(const bstrPrinterName: WideString; var hdc: _RemotableHandle): HResult; stdcall;
  end;

{ capture an IHTMLImgElement from a page. bmp must be created&freed by the caller.(content maybe modified) }
function CaptureHtmlImg(const AImg: IHTMLIMGElement; var Bmp: TBitmap): Boolean;
begin
  Result := False;
  if not Supports(AImg, IHTMLElementRender) then Exit;
//  LogFmt('width=%d height=%d', [AImg.width, AImg.height]);
  if (AImg.width = 0) or (AImg.height = 0) then Exit;
  Bmp.Width := AImg.width;
  Bmp.Height := AImg.height;
  Result := Succeeded((AImg as IHTMLElementRender).DrawToDC(Bmp.Canvas.Handle));
end;

function SetSelectOption(ADoc: IHTMLDocument2; ASelectId: string; AOption: string): Boolean;
var
  Elem: IHTMLElement;
  Select: IHTMLSelectElement;
  I: Integer;
  Option: IHTMLOptionElement;
begin
  Result := False;
  Elem := ADoc.all.item(AselectId, 0) as IHTMLElement;
  if Elem = nil then Exit;
  if not Supports(Elem, IHTMLSelectElement) then Exit;
  Select := Elem as IHTMLSelectElement;
  for I := 0 to Select.length - 1 do
  begin
    Elem := Select.item(I, 0) as IHTMLElement;
    if not Supports(Elem, IHTMLOptionElement) then Continue;
    Option := Elem as IHTMLOptionElement;
    if not SameText(Trim(Option.text), AOption) then Continue;
    Select.selectedIndex := I;
    SafeFireEvent(Select, 'onchange', []);
    Result := True;
    Break;
  end;
end;

function SetSelectOptionByValue(ADoc: IHTMLDocument2; ASelectId: string; AValue: string): Boolean;
var
  Elem: IHTMLElement;
  Select: IHTMLSelectElement;
  I: Integer;
  Option: IHTMLOptionElement;
begin
  Result := False;
  Elem := ADoc.all.item(AselectId, 0) as IHTMLElement;
  if Elem = nil then Exit;
  if not Supports(Elem, IHTMLSelectElement) then Exit;
  Select := Elem as IHTMLSelectElement;
  for I := 0 to Select.length - 1 do
  begin
    Elem := Select.item(I, 0) as IHTMLElement;
    if not Supports(Elem, IHTMLOptionElement) then Continue;
    Option := Elem as IHTMLOptionElement;
    if not SameText(Trim(Option.value), AValue) then Continue;
    Select.selectedIndex := I;
    SafeFireEvent(Select, 'onchange', []);
    Result := True;
    Break;
  end;
end;

function SetInputTagText(ADoc: IHTMLDocument2; AInputId: string; AText: string): Boolean;
var
  Elem: IHTMLElement;
  I: Integer;
begin
  Result := False;
  I := 0;
  while True do
  begin
    Elem := ADoc.all.item(AInputId, I) as IHTMLElement;
    if Elem = nil then Exit;
    if Supports(Elem, IHTMLInputElement) then
    begin
      if SameText((Elem as IHTMLInputElement).type_, 'text')
        or SameText((Elem as IHTMLInputElement).type_, 'password')
        or SameText((Elem as IHTMLInputElement).type_, 'hidden') then
      begin
        (Elem as IHTMLInputElement).value := AText;
        Result := True;
        Exit;
      end;
    end;
    Inc(I);
  end;
end;

function CalcElementClientRect(AElem: IHTMLElement): TRect;
var
  CR: IHTMLRectCollection;
  I: Integer;                    
  R: IHTMLRect;
  V: OleVariant;
begin
  Result := Rect(0, 0, 0, 0);
  CR := (AElem as IHTMLElement2).getClientRects;
  if CR.length = 0 then Exit;
  V := 0;
  R := IUnknown(CR.item(V)) as IHTMLRect;
  Result := Rect(R.left, R.top, R.right, R.bottom);
  for I := 1 to CR.length - 1 do
  begin
    V := I;
    R := IUnknown(CR.item(V)) as IHTMLRect;
    if R.left < Result.Left then Result.Left := R.left;
    if R.top < Result.Top then Result.Top := R.top;
    if R.right > Result.Right then Result.Right := R.right;
    if R.bottom > Result.Bottom then Result.Bottom := R.bottom;
  end;
end;

function CalcElementClientRect2(AElem, ABody: IHTMLElement): TRect;
var
  W, H: Integer;
begin
  W := AElem.offsetWidth;
  H := AElem.offsetHeight;
  SetRect(Result, (ABody as IHTMLElement2).clientLeft, (ABody as IHTMLElement2).clientTop, 0, 0);
  while AElem <> nil do
  begin
    Inc(Result.Left, AElem.offsetLeft);
    Inc(Result.Top, AElem.offsetTop);
    AElem := AElem.offsetParent;
    if AElem = ABody then Break;
  end;
  Result.Right := Result.Left + W;
  Result.Bottom := Result.Top + H;
end;

{ TCallUntilDoneObj }

constructor TCallUntilDoneObj.Create(AWB: TWebBrowser;
  ACallback: TNavigateAndCallFunc; ATimeoutFunc: TNotifyEvent);
begin
  FWB := AWB;
  FCallback := ACallback;
  FTimer := TTimer.Create(nil);
  FTimer.Interval := 2000;
  FTimer.Enabled := True;
  FTimer.OnTimer := OnTimer;
  FTimeCount := 40;
end;

destructor TCallUntilDoneObj.Destroy;
begin
  FreeAndNil(FTimer);
  l_TempObjList.Remove(Self);
  inherited;
end;

procedure TCallUntilDoneObj.OnTimer(Sender: TObject);
var
  Done: Boolean;
begin
  FTimer.Enabled := False;
  if Assigned(FCallback) then
  begin
    try
      FCallback(FWB, Done);
    except
      Done := True;
    end;
    if Done then
    begin
      Self.Free;
      Exit;
    end;
  end;
  Dec(FTimeCount);
  if FTimeCount <= 0 then
  begin
    if Assigned(FTimeoutFunc) then
    begin
      try
        FTimeoutFunc(Self);
      except
        ;
      end;
    end;
    Self.Free;
    Exit;
  end;
  FTimer.Enabled := True;
end;

{ TCSVGridParser }

constructor TCSVGridParser.Create(AFileName: string);
begin
  if not FileExists(AFileName) then
    raise EHtmlParseException.CreateFmt('文件%s不存在！', [AFileName]);
  FCols := TStringList.Create;
  FLines := TStringList.Create;
  FLines.LoadFromFile(AFileName);
  FLineSplitted := nil;
  FCurLine := 0;
  FLastLine := -1;
end;

destructor TCSVGridParser.Destroy;
begin
  if Assigned(FLineSplitted) then FreeAndNil(FLineSplitted);
  FreeAndNil(FLines);
  FreeAndNil(FCols);
  inherited;
end;

function TCSVGridParser.EOT: Boolean;
begin
  Result := FCurLine > FLastLine;
end;

function TCSVGridParser.GetCell(AColName: string): string;
var
  I: Integer;
begin
  I := FCols.IndexOf(AColName);
  if I < 0 then raise EHtmlParseException.CreateFmt('column %s not found', [AColName]);
  LineSplittedRequired;
  if I >= FLineSplitted.Count then raise EHtmlParseException.Create('index out of range');
  Result := RefineCell(Trim(FLineSplitted[I]));
end;

function TCSVGridParser.GetCellDef(AColName: string;
  const ADef: string): string;
var
  I: Integer;
begin
  Result := ADef;
  I := FCols.IndexOf(AColName);
  if I < 0 then Exit;
  LineSplittedRequired;
  if I >= FLineSplitted.Count then Exit;
  Result := RefineCell(Trim(FLineSplitted[I]));
end;

function TCSVGridParser.NextRow: Boolean;
begin
  Result := False;
  repeat
    if Assigned(FLineSplitted) then FreeAndNil(FLineSplitted);
    Inc(FCurLine);
    if FCurLine > FLastLine then Exit;
    LineSplittedRequired;
    if FLineSplitted.Count = FCols.Count then
    begin
      Result := True;
      Exit;
    end;
  until False;
end;

function TCSVGridParser.ParseHeader(
  ARequiredCols: array of string): Boolean;
var
  I, J: Integer;
  SL: TStringList;
  S: string;
begin
  if Assigned(FLineSplitted) then FreeAndNil(FLineSplitted);
  FCurLine := 0;
  Result := False;
  // refine texts
  for I := 0 to FLines.Count - 1 do
  begin
    S := Trim(FLines[I]);
    while (S <> '') and (S[Length(S)] = ',') do S := Copy(S, 1, Length(S) - 1);
    FLines[I] := S;
  end;
  for I := 0 to FLines.Count - 1 do
  begin
    FCols.CommaText := FLines[I];
    if FCols.Count < Length(ARequiredCols) then Continue;
    Result := True;
    for J := Low(ARequiredCols) to High(ARequiredCols) do
    begin
      if FCols.IndexOf(ARequiredCols[J]) < 0 then
      begin
        Result := False;
        Break;
      end;
    end;
    if not Result then Continue;
    FCurLine := I + 1;
    Break;
  end;
  if not Result then Exit;
  SL := TStringList.Create;
  try
    FLastLine := -1;
    for I := FLines.Count - 1 downto 0 do
    begin
      SL.CommaText := FLines[I];
      if SL.Count = FCols.Count then
      begin
        FLastLine := I;
        Break;
      end;
    end;
    for I := FCurLine to FLastLine do
    begin
      SL.CommaText := FLines[I];
      if SL.Count = FCols.Count then
      begin
        FCurLine := I;
        Break;
      end;
    end;
  finally
    FreeAndNil(SL);
  end;
end;

procedure TCSVGridParser.LineSplittedRequired;
begin
  if not Assigned(FLineSplitted) then
  begin
    FLineSplitted := TStringList.Create;
    FLineSplitted.CommaText := FLines[FCurLine];
  end;
end;

function TCSVGridParser.RefineCell(S: string): string;
var
  P: PChar;
begin
  if S <> '' then
  begin
    if (S[1] = '''') and (S[Length(S)] = '''') then
    begin
      P := PChar(S);
      Result := Trim(AnsiExtractQuotedStr(P, ''''));
      Exit;
    end
    else if (S[1] = '"') and (S[Length(S)] = '"') then
    begin
      P := PChar(S);
      Result := Trim(AnsiExtractQuotedStr(P, '"'));
      Exit;
    end;
  end;
  Result := S;
end;

function TCSVGridParser.TextExists(AText: string): Boolean;
begin
  Result := Pos(AText, FLines.Text) > 0;
end;

{ TCommonExchangeFileParserBase }

function TCommonExchangeFileParserBase._AddRef: Integer;
begin
  Result := 0;
end;

function TCommonExchangeFileParserBase._Release: Integer;
begin
  Result := 0;
end;

function TCommonExchangeFileParserBase.QueryInterface(const IID: TGUID;
  out Obj): HResult;
begin
  Result := E_NOTIMPL;
end;

{ THorizTableParser }

function THorizTableParser._AddRef: Integer;
begin
  Result := 0;
end;

function THorizTableParser._Release: Integer;
begin
  Result := 0;
end;

function THorizTableParser.QueryInterface(const IID: TGUID;
  out Obj): HResult;
begin
  Result := E_NOTIMPL;
end;

constructor THorizTableParser.Create(ATable: IHTMLTable);
var
  TDs: IHTMLElementCollection;
  I: Integer;
  Cell: IHTMLElement;
begin
  FCells := TStringList.Create;
  FCells.Sorted := False;
  TDs := ((ATable as IHTMLElement).all as IHTMLElementCollection).tags('td') as IHTMLElementCollection;
  for I := 0 to TDs.length - 1 do
  begin
    Cell := TDs.item(I, 0) as IHTMLElement;
    FCells.Add(Trim(Cell.innerText));
  end;
end;

destructor THorizTableParser.Destroy;
begin
  FreeAndNil(FCells);
  inherited;
end;

function THorizTableParser.GetCell(AColName: string): string;
var
  I: Integer;
begin
  I := FCells.IndexOf(AColName);
  if I < 0 then raise EHtmlParseException.CreateFmt('col %s does not exists!', [AColName]);
  if I + 1 >= FCells.Count then raise EHtmlParseException.CreateFmt('col %s has no value!', [AColName]);
  Result := FCells[I + 1];  
end;

function THorizTableParser.GetCellDef(AColName: string;
  const ADef: string): string;
var
  I: Integer;
begin
  I := FCells.IndexOf(AColName);
  if I < 0 then Result := ADef
  else if I + 1 >= FCells.Count then Result := ADef
  else Result := FCells[I + 1];
end;

procedure THorizTableParser.SetFirstColumn(const ACol: string);
var
  I: Integer;
begin
  I := FCells.IndexOf(ACol);
  if I < 0 then Exit;
  while I > 0 do
  begin
    FCells.Delete(0);
    Dec(I);
  end;
end;

function THorizTableParser.GetCellHtmlDef(AColName: string;
  const ADef: string): string;
begin
  Result := '';
end;

{ TCrossTableParser }

constructor TCrossTableParser.Create(ATable: IHTMLTable);
begin
  FTable := ATable;
  FColHeader := TStringList.Create;
  FRowHeader := TStringList.Create;
end;

destructor TCrossTableParser.Destroy;
begin
  FTable := nil;
  FreeAndNil(FColHeader);
  FreeAndNil(FRowHeader);
  inherited;
end;

function TCrossTableParser.GetCellDef(AColName, ARowName: string;
  const ADef: string): string;
var
  RowIndex, ColIndex: Integer;
  Row: IHTMLTableRow;
  Cell: IHTMLTableCell;
begin
  RowIndex := FRowHeader.IndexOf(ARowName);
  ColIndex := FColHeader.IndexOf(AColName);
  if (RowIndex < 0) or (ColIndex < 0) then Result := ADef
  else
  begin
    Row := FTable.rows.item(RowIndex, 0) as IHTMLTableRow;
    if ColIndex >= Row.cells.length then Result := ADef
    else
    begin
      Cell := Row.cells.item(ColIndex, 0) as IHTMLTableCell;
      Result := Trim((Cell as IHTMLElement).innerText);
    end;
  end;
end;

function TCrossTableParser.ParseHeader(ColHeader: array of string): Boolean;
var
  I, J, K, L, M: Integer;
  Row: IHTMLTableRow;
  Flags: array of Boolean;
  S: string;
begin
  FColHeader.Clear;
  FRowHeader.Clear;
  Result := False;
  // parse col header(a row)
  SetLength(Flags, Length(ColHeader));
  try
    for I := 0 to FTable.rows.length - 1 do
    begin
      Row := FTable.rows.item(I, 0) as IHTMLTableRow;
      if Row.cells.length < Length(ColHeader) then Continue;
      for J := High(Flags) downto 0 do Flags[J] := False;
      K := Length(ColHeader);
      for J := 0 to Row.cells.length - 1 do
      begin
        S := Trim((Row.cells.item(J, 0) as IHTMLElement).innerText);
        for L := 0 to High(ColHeader) do
        begin
          if Flags[L] then Continue;
          if SameText(S, ColHeader[L]) then
          begin
            Flags[L] := True;
            Dec(K);
            if K = 0 then
            begin
              for M := 0 to Row.cells.length - 1 do FColHeader.Add(Trim((Row.cells.item(M, 0) as IHTMLElement).innerText));
              Abort;
            end;
          end;
        end;
      end;
    end;
  except
    Result := True; // to catch the abort
  end;
  // parse row header(a column)
  for I := 0 to FTable.rows.length - 1 do
  begin
    Row := FTable.rows.item(I, 0) as IHTMLTableRow;
    if Row.cells.length = 0 then S := ''
    else S := Trim((Row.cells.item(0, 0) as IHTMLElement).innerText);
    FRowHeader.Add(S);
  end;
end;

function _IEPopupDialogHookProp(nCode: Integer; wp: WPARAM; lp: LPARAM): LRESULT; stdcall;
var
  pcwp: PCWPStruct;
  ClsName, WndName: string;
begin
  Result := CallNextHookEx(l_OldCallWndProcHook, nCode, wp, lp);
  pcwp := PCWPStruct(lp);
  if (pcwp.message = WM_SHOWWINDOW) and (pcwp.wParam = 1) and IsWindow(pcwp.hwnd) then
  begin
    SetLength(ClsName, 64);
    SetLength(ClsName, GetClassName(pcwp.hwnd, PChar(ClsName), 64));
    if SameText(ClsName, 'Internet Explorer_TridentDlgFrame') then
    begin
      { 出现"脚本错误"对话框.它是一个网页对话框,在显示时页面并没有加载完成,因此
        此时找不到"是"按钮. 这里首先把它的大小设为0x0,然后等待它加载完成. }
      if Assigned(l_IEPopupDialogWatcher) then
      begin
        MoveWindow(pcwp.hwnd, 0, 0, 0, 0, False);
        PostMessage(l_IEPopupDialogWatcher.Handle, WM_PROCESS_YESNODLG, pcwp.hwnd, 0);
      end;
    end
    else if SameText(ClsName, '#32770') then
    begin
      SetLength(WndName, 32);
      SetLength(WndName, GetWindowText(pcwp.hwnd, PChar(WndName), 32));
      if SameText(WndName, '错误') then
      begin
        if (GetDlgItem(pcwp.hwnd, IDYES) <> 0) and (GetDlgItem(pcwp.hwnd, IDNO) <> 0) then
        begin
          { 出现脚本错误调试对话框。是否调试->选否 }
          MoveWindow(pcwp.hwnd, 0, 0, 0, 0, False);
          PostMessage(pcwp.hwnd, WM_COMMAND, IDNO, 0);
        end;
      end
      else if SameText(WndName, 'Microsoft Internet Explorer') then
      begin
        if FindWindowEx(pcwp.hwnd, 0, nil, '当前安全设置禁止运行该页中的 ActiveX 控件。因此，该页可能无法正常显示。') <> 0 then
        begin
          { 出现安全设置对话框。关闭 }
          MoveWindow(pcwp.hwnd, 0, 0, 0, 0, False);
          PostMessage(pcwp.hwnd, WM_COMMAND, IDCANCEL, 0);
        end;
      end;
    end;
  end;
end;

{ TIEPopupDialogWatcher }

constructor TIEPopupDialogWatcher.Create;
begin
  FHandle := Classes.AllocateHWnd(WndProc);
  if l_OldCallWndProcHook = 0 then
    l_OldCallWndProcHook := SetWindowsHookEx(WH_CALLWNDPROC, _IEPopupDialogHookProp, 0, GetCurrentThreadId);
end;

destructor TIEPopupDialogWatcher.Destroy;
begin
  if l_OldCallWndProcHook <> 0 then UnhookWindowsHookEx(l_OldCallWndProcHook);
  Classes.DeallocateHWnd(FHandle);
  inherited;
end;

procedure TIEPopupDialogWatcher.OnProcessYesNoDlg(var Msg: TMessage);
var
  Doc: IHTMLDocument2;
  Elem: IHTMLElement;
begin
  if not IsWindow(Msg.WParam) then Exit;
  Doc := GetIHTMLDocumentFromHWND(Msg.WParam);
  if Doc <> nil then
  begin
    Elem := Doc.all.item('btnYes', 0) as IHTMLElement;
    if Elem <> nil then
    begin
      Elem.click;
    end;
  end;
  PostMessage(Handle, WM_PROCESS_YESNODLG, Msg.WParam, 0);
end;

procedure TIEPopupDialogWatcher.WndProc(var Msg: TMessage);
begin
  if Msg.Msg < WM_USER then Msg.Result := DefWindowProc(Handle, Msg.Msg, Msg.WParam, Msg.LParam)
  else Dispatch(Msg);
end;

{ TTempObjList }

procedure TTempObjList.Add(AObj: TObject);
begin
  inherited Add(AObj);
end;

constructor TTempObjList.Create;
begin
  FHandle := Classes.AllocateHWnd(WndProc);
end;

destructor TTempObjList.Destroy;
begin
  Classes.DeallocateHWnd(FHandle);
  while Self.Count > 0 do
  begin
    TObject(Items[0]).Free;
  end;
  inherited;
end;

procedure TTempObjList.OnCallerDestroyed(ACaller: TObject);
type
  TNotifyEventRec = packed record
    PFunc: Pointer;
    PObj: Pointer;
  end;
var
  I: Integer;
  Obj: TObject;
  Rec: TNotifyEventRec;
  NeedRemove: Boolean;
begin
  for I := Self.Count - 1 downto 0 do
  begin
    Obj := TObject(Items[I]);
    NeedRemove := False;
    if Obj is TCallUntilDoneObj then
    begin
      CopyMemory(@Rec, @@TCallUntilDoneObj(Obj).FCallback, 8);
      if Rec.PObj = ACaller then NeedRemove := True;
      CopyMemory(@Rec, @@TCallUntilDoneObj(Obj).FTimeoutFunc, 8);
      if Rec.PObj = ACaller then NeedRemove := True;
    end
    else if Obj is TNavigateAndCallObj then
    begin
      CopyMemory(@Rec, @@TNavigateAndCallObj(Obj).FCallback, 8);
      if Rec.PObj = ACaller then NeedRemove := True;
      CopyMemory(@Rec, @@TNavigateAndCallObj(Obj).FOnTimeout, 8);
      if Rec.PObj = ACaller then NeedRemove := True;
    end;
    if NeedRemove then TObject(Items[I]).Free;
  end;
end;

procedure TTempObjList.Remove(AObj: TObject);
begin
  inherited Remove(AObj);
end;

procedure UnregisterCallerObject(AObject: TObject);
begin
  l_TempObjList.OnCallerDestroyed(AObject);
end;

function SafeFireEvent(ADisp: IDispatch; AEventName: string; Params: array of const): Boolean;
type
  PVarArg = ^TVarArg;
  TVarArg = array[0..3] of DWORD;
var
  Name: WideString;
  DispID: Integer;
  DispParams: TDispParams;
  I: Integer;
  Args: array[0..31] of OleVariant;
  InvokeResult: OleVariant;
  Rec: TVarRec;
  V: OleVariant;
  TmpDisp: IUnknown;
  ExcepInfo: TExcepInfo;
begin
  Result := False;
  Name := AEventName;
  if Succeeded(ADisp.GetIDsOfNames(GUID_NULL, @Name, 1, 0, @DispID)) then
  begin
    DispParams.cArgs := Length(Params);
    DispParams.cNamedArgs := 0;
    for I := 0 to High(Params) do
    begin
      Rec := Params[I];
      case Rec.VType of
        vtInteger:    V := Rec.VInteger;
        vtBoolean:    V := Rec.VBoolean;
        vtChar:       V := Rec.VChar;
        vtExtended:   V := Rec.VExtended^;
        vtString:     V := Rec.VString^;
        vtPointer:    raise Exception.Create('pointer type not supported');
        vtPChar:      V := string(Rec.VPChar);
        vtObject:     begin
                        if Rec.VObject.GetInterface(IUnknown, TmpDisp) then V := TmpDisp
                        else raise Exception.Create('object type not supported');
                      end;

        vtWideChar:   V := Rec.VWideChar;
        vtPWideChar:  V := WideString(Rec.VPWideChar);
        vtAnsiString: V := string(Rec.VAnsiString^);
        vtCurrency:   V := Rec.VCurrency^;
        vtVariant:    V := Rec.VVariant^;
        vtInterface:  V := IUnknown(Rec.VInterface^);
        vtWideString: V := WideString(Rec.VWideString^);
        vtInt64:      V := Rec.VInt64^
        else raise Exception.Create('type not supported!');
      end;
      Args[I] := V;
    end;
    DispParams.rgvarg := @Args;
    Result := Succeeded(ADisp.Invoke(DispID, GUID_NULL, 0, DISPATCH_METHOD, DispParams, @InvokeResult, @ExcepInfo, nil));
  end;
end;

procedure TTempObjList.WndProc(var Msg: TMessage);
var
  Obj: Pointer;
  Index: Integer;
begin
  if Msg.Msg = WM_REMOVE_OBJ then
  begin
    Obj := Pointer(Msg.LParam);
    Index := IndexOf(Obj);
    if Index >= 0 then
    begin
      Self.Delete(Index);
      TObject(Obj).Free;
    end;
  end else Msg.Result := DefWindowProc(FHandle, Msg.Msg, Msg.WParam, Msg.LParam);
end;

procedure FramePostData(AWB: TWebBrowser; AUrl: string; APostData: string; AFrameId: string);
var
  PostData, Header: OleVariant;
  Flags: OleVariant;
  FrameName: OleVariant;
  P: Pointer;
begin
  Flags := 0;
  PostData := VarArrayCreate([0, Length(APostData) - 1], varByte);
  P := VarArrayLock(PostData);
  try
    Move( PChar(APostData)^, P^, Length(APostData) );
  finally
    VarArrayUnlock(PostData);
  end;
  Header := 'Content-Type: application/x-www-form-urlencodedrn';
  FrameName := AFrameId;
  AWB.Navigate(AUrl, Flags, FrameName, PostData, Header);
end;

{ ******************** SetCustomActiveXControlText ******************** }
type
  PImageImportDescriptor = ^TImageImportDescriptor;
  TImageImportDescriptor = packed record
    OriginalFirstThunk: DWORD;  // or Characteristics: DWORD
    TimeDateStamp: DWORD;
    ForwarderChain: DWORD;
    Name: DWORD;
    FirstThunk: DWORD;
  end;
  PImageChunkData = ^TImageChunkData;
  TImageChunkData = packed record
    case Integer of
      0: ( ForwarderString: DWORD );
      1: ( Func: DWORD );
      2: ( Ordinal: DWORD );
      3: ( AddressOfData: DWORD );
  end;
  PImageImportByName = ^TImageImportByName;
  TImageImportByName = packed record
    Hint: Word;
    Name: array[0..0] of Byte;
  end;

type
  PHookRec = ^THookRec;
  THookRec = packed record
    OldFunc: Pointer;
    NewFunc: Pointer;
  end;

procedure HookApiInMod(ImageBase: Cardinal; DllName: PChar; ApiName: PChar; PHook: PHookRec);
var
  pidh: PImageDosHeader;
  pinh: PImageNtHeaders;
  pSymbolTable: PIMAGEDATADIRECTORY;
  piid: PIMAGEIMPORTDESCRIPTOR;
  written, oldAccess: DWORD;
  pProtoFill: Pointer;
  Loaded: HMODULE;
  pCode: ^Pointer;
begin
  if ImageBase = 0 then Exit;
  Loaded := LoadLibrary(DllName);
  pProtoFill := GetProcAddress(Loaded, ApiName);

  pidh := PImageDosHeader(ImageBase);
  pinh := PImageNtHeaders(DWORD(ImageBase) + Cardinal(pidh^._lfanew));
  pSymbolTable := @pinh^.OptionalHeader.DataDirectory[1];
  piid := PImageImportDescriptor(DWORD(ImageBase) + pSymbolTable^.VirtualAddress);

  while piid^.Name <> 0 do
  begin
    pCode := Pointer(dword(ImageBase) + piid^.FirstThunk);
    while pCode^ <> nil do
    begin
      if (pCode^ = pProtoFill) then
      begin
        PHook^.OldFunc := pCode^;
        VirtualProtect(pCode, SizeOf(DWORD), PAGE_WRITECOPY, oldAccess);
        WriteProcessMemory(GetCurrentProcess(), pCode, @PHook^.NewFunc, SizeOf(DWORD), written);
        VirtualProtect(pCode, SizeOf(DWORD), oldAccess, oldAccess);
      end;
      pCode := Pointer(dword(pCode) + 4);
    end;
    piid := Pointer(dword(piid) + 20);
  end;
end;

procedure UnHookApiInMod(ImageBase: Cardinal; DllName: PChar; ApiName: PChar; const PHook: PHookRec);
var
  pidh: PImageDosHeader;
  pinh: PImageNtHeaders;
  pSymbolTable: PIMAGEDATADIRECTORY;
  piid: PIMAGEIMPORTDESCRIPTOR;
  written, oldAccess: DWORD;
  pCode: ^Pointer;
begin
  if ImageBase = 0 then Exit;

  pidh := PImageDosHeader(ImageBase);
  pinh := PImageNtHeaders(DWORD(ImageBase) + Cardinal(pidh^._lfanew));
  pSymbolTable := @pinh^.OptionalHeader.DataDirectory[1];
  piid := PImageImportDescriptor(DWORD(ImageBase) + pSymbolTable^.VirtualAddress);

  while piid^.Name <> 0 do
  begin
    pCode := Pointer(dword(ImageBase) + piid^.FirstThunk);
    while pCode^ <> nil do
    begin
      if (pCode^ = pHook^.NewFunc) then
      begin
        VirtualProtect(pCode, SizeOf(DWORD), PAGE_WRITECOPY, oldAccess);
        WriteProcessMemory(GetCurrentProcess(), pCode, @PHook^.OldFunc, SizeOf(DWORD), written);
        VirtualProtect(pCode, SizeOf(DWORD), oldAccess, oldAccess);
      end;
      pCode := Pointer(dword(pCode) + 4);
    end;
    piid := Pointer(dword(piid) + 20);
  end;
end;

type

  PHookProcInstanceRec = ^THookProcInstanceRec;
  THookProcInstanceRec = packed record
    Code:     Byte;       { $E8  CALL NEAR PTR }
    Offset:   Integer;    { offset }
    Obj:      Pointer;    { pointer of this }
    Code1:    Byte;       { $59  pop ecx }
    Code2:    Byte;       { $E9  JMP NEAR PTR }
    StdProc:  Integer;    { offset stdproc }
  end;

  TGetFocusContext = class
  private
    FCode: PHookProcInstanceRec;
  public
    ForceFocus: HWND;
    HookRec: THookRec;
    function GetFocus: HWND; stdcall;
    function NewProcPtr: Pointer;
    constructor Create;
    destructor Destroy; override;
  end;

type
  TGetFocusFunc = function: HWND; stdcall;

function _StdGetFocusFunc: HWND; stdcall; assembler;
asm
  MOV   ECX, [ECX]
  PUSH  ECX
  CALL  TGetFocusContext.GetFocus
end;

function _SimpleSimulateKeyStrikes(const S: string): Boolean;
var
  Code: Short;
  I: Integer;
  Ch: Char;
  CapslockDown: Boolean;
  NeedShift: Boolean;

  procedure EnsureKeyEvent(ACode: Byte; ADown: Boolean);
  begin
    if ADown then keybd_event(ACode, 1, 0, 0)
    else keybd_event(ACode, 1, KEYEVENTF_KEYUP, 0);
    Sleep(20);
  end;

begin
  CapslockDown := (GetKeyState(VK_CAPITAL) and 1) <> 0;
  for I := 1 to Length(S) do
  begin
    Code := VkKeyScan(S[I]);
    Ch := Char(Code and $FF);
    NeedShift := (Code and $100) <> 0;
    if Ch in ['A'..'Z'] then
    begin
      if CapslockDown then NeedShift := not NeedShift;
    end;
    if NeedShift then  // shift is down
    begin
      EnsureKeyEvent(VK_SHIFT, True);
      EnsureKeyEvent(Code, True);
      EnsureKeyEvent(Code, False);
      EnsureKeyEvent(VK_SHIFT, False);
    end
    else
    begin
      EnsureKeyEvent(Code, True);
      EnsureKeyEvent(Code, False);
    end;
  end;
  Result := True;
end;

const
  PROP_OLDWNDPROC = 'BM_OLD_WNDPROC';

var
  l_SCACTEvent: TEvent;

type
  TStdWndProc = function (hwnd: HWND; uMsg: UINT; wParam: WPARAM; lParam: LPARAM): LRESULT; stdcall;

function StdDisableFocusChangeWndProc(H: HWND; uMsg: UINT; wp: WPARAM; lp: LPARAM): LRESULT; stdcall;
var
  OldWndProc: TStdWndProc;
begin
  OldWndProc := TStdWndProc(GetProp(H, PROP_OLDWNDPROC));
  if ((uMsg = WM_SETFOCUS) or (uMsg = WM_KILLFOCUS)) and (HWND(lp) <> H) then
  begin
    Result := 0;
//    Log('WM_SETFOCUS/WM_KILLFOCUS eat');
  end
  else
  begin
    if (uMsg = WM_SETFOCUS) or (uMsg = WM_KILLFOCUS) then
    begin
//      Log('WM_SETFOCUS/WM_KILLFOCUS pass');
    end;
    Result := OldWndProc(H, uMsg, wp, lp);
  end;
end;

procedure SetCustomActiveXControlText(AElem: IHTMLElement; AText: string; ADllName: string);
var
  H: HWND;
  Ctx: TGetFocusContext;
  OldWndProc: LongInt;
  Msg: tagMsg;
begin
  while True do
  begin
    if l_SCACTEvent.WaitFor(100) = wrSignaled then Break;
    Application.ProcessMessages;
  end;
  try
    (AElem as IOleWindow).GetWindow(H);
    Ctx := TGetFocusContext.Create;
    try
      Ctx.HookRec.OldFunc := nil;
      Ctx.HookRec.NewFunc := Ctx.NewProcPtr;
      if GetModuleHandle(PChar(ADllName)) = 0 then raise Exception.CreateFmt('%s not loaded', [ADllName]);
      // hook the getfocus function called by the control
      HookApiInMod( GetModuleHandle(PChar(ADllName)), 'user32.dll', 'GetFocus', @Ctx.HookRec );
      Ctx.ForceFocus := H;
      // send this message so the control will set the LL_Keyboard hook
      OldWndProc := GetWindowLong(H, GWL_WNDPROC);
      SetProp(H, PROP_OLDWNDPROC, OldWndProc);
      SetWindowLong(H, GWL_WNDPROC, Integer(@StdDisableFocusChangeWndProc));

      SendMessage(H, WM_SETFOCUS, 0, H);
      // now simulate keystrikes, all keys will go to the control & because getfocus returns itself, it will accept all keys
      _SimpleSimulateKeyStrikes(AText);

      // process all keyboard & other messages generated during the keybd_event
      while PeekMessage(Msg, H, WM_KEYFIRST, WM_USER, PM_REMOVE) do
      begin
        TranslateMessage(Msg);
        DispatchMessage(Msg);
      end;

      Ctx.ForceFocus := 0;
      SetWindowLong(H, GWL_WNDPROC, OldWndProc);
      RemoveProp(H, PROP_OLDWNDPROC);

      // send this message so the control will unhook the LL_Keyboard hook
      SendMessage(H, WM_KILLFOCUS, 0, H);

//      Log('simulate done');
      // unhook the function
      UnHookApiInMod( GetModuleHandle(PChar(ADllName)), 'user32.dll', 'GetFocus', @Ctx.HookRec );
    finally
      FreeAndNil(Ctx);
    end;
  finally
    l_SCACTEvent.SetEvent;
  end;
end;

{ TGetFocusContext }

function CalcJmpOffset(Src, Dest: Pointer): Longint;
begin
  Result := Longint(Dest) - (Longint(Src) + 5);
end;

constructor TGetFocusContext.Create;
begin
  FCode := VirtualAlloc(nil, SizeOf(THookProcInstanceRec), MEM_COMMIT, PAGE_EXECUTE_READWRITE);
  FCode.Code := $E8;
  FCode.Offset := CalcJmpOffset(@FCode.Code, @FCode.Code1);
  FCode.Obj := Self;
  FCode.Code1 := $59;
  FCode.Code2 := $E9;
  FCode.StdProc := CalcJmpOffset(@FCode.Code2, @_StdGetFocusFunc);
end;

destructor TGetFocusContext.Destroy;
begin
  VirtualFree(FCode, 0, MEM_RELEASE);
  inherited;
end;

function TGetFocusContext.GetFocus: HWND;
begin
  if ForceFocus = 0 then Result := TGetFocusFunc(HookRec.OldFunc)()
  else Result := ForceFocus;
//  LogFmt('GetFocus returns:%X', [Result]);
end;

function TGetFocusContext.NewProcPtr: Pointer;
begin
  Result := FCode;
end;

procedure DisableCurrentThreadIME;
type
  TImmDisableIMEFunc = function (AThreadId: Cardinal): LongBool; stdcall;
var
  _ImmDisableIME: TImmDisableIMEFunc;
begin
  _ImmDisableIME := GetProcAddress( GetModuleHandle('imm32.dll'), 'ImmDisableIME' );
  if @_ImmDisableIME <> nil then _ImmDisableIME(GetCurrentThreadId);
end;

initialization

  l_TempObjList := TTempObjList.Create;
  l_IEPopupDialogWatcher := TIEPopupDialogWatcher.Create;
  l_SCACTEvent := TEvent.Create(nil, False, True, '');

finalization

  FreeAndNil(l_SCACTEvent);
  l_TempObjList.Free;
  l_TempObjList := nil;
  FreeAndNil(l_IEPopupDialogWatcher);

end.
