unit WBrowserEditor;
{-----------------------------------------------------------------------------
  这是一个拿 Ie 浏览器来作为富文本编辑器即 Html 编辑器的控件

  用 IE 浏览器来做编辑界面：
  1. 拖一个 TWebBrowser 下来。
  2. 打开这个浏览器，执行 WebBrowser1.Navigate('about:blank');
  3. 在 TWebBrowser 的事件：WebBrowser1NavigateComplete2 里，取得 IHTMLDocument2 接口
  4. 将浏览器设置为编辑模式，即将取得的 IHTMLDocument2 接口的 DesignMode 设置为 'On' 字符串
  5. 要设置选中的字符串的字体、颜色等，调用 IHTMLDocument2 接口的 execCommand 方法。
  6. execCommand 方法里要改被选中的字体大小，则方法名就是字符串 'FontSize'，如果要改为粗体，方法名就是 'Bold'
  7. 输入的时候，这个对象一定要 SetFocus 否则不能输入 Enter 回车。奇怪的是鼠标点进去可以输入的时候，并不是 SetFocus 的。
  8. 作为编辑器，对插入的图像的处理，分为两种模式：
     A. 把图像的数据用 Base64 编码后直接夹在正文数据里传输给对方
     B. 只把正文里图像的链接描述字符串送给对方。

  9. MSHTML 对表格不支持。需要自己写。
  
     详细的方法名和参数，参考：
     http://msdn.microsoft.com/workshop/browser/mshtml/reference/ifaces/document2/execcommand.asp
     http://msdn.microsoft.com/workshop/author/dhtml/reference/commandids.asp
     http://msdn.microsoft.com/workshop/browser/mshtml/reference/ifaces/document2/document2.asp
     http://msdn.microsoft.com/workshop/browser/editing/mshtmleditor.asp

     这里有个代码可以下载来看一下：
     http://www.itwriting.com/htmleditor/index.php

  pcplayer. 2006-2-15

  今天升级为支持 delphi 2010 的控件。pcplayer 2010-5-31
------------------------------------------------------------------------------}

interface

uses
  Winapi.Windows,  Winapi.Messages, Winapi.ActiveX,
  System.SysUtils, System.Variants, System.Classes, System.StrUtils, System.TypInfo,
  Vcl.Controls, Vcl.Consts, Vcl.Graphics, Vcl.Forms,
  MSHTML, MSHTMLEvents, Vcl.OleCtrls, SHDocVw,

  MiMeEncodePic{编码图片用}{, LjnMailMessage{解码用};

type
  //事件
  TMouseDownEvent = procedure(Sender:Tobject;Button:TMouseButton;Shift:TShiftState;x,y:integer) of object;
  TMouseUpEvent = procedure(Sender:Tobject;Button:TMouseButton;Shift:TShiftState;x,y:integer) of object;
  TMouseMoveEvent = procedure(Sender:Tobject;Shift:TShiftState;x,y:integer) of object;
  TEnterEvent = procedure(Sender:Tobject) of object;
  TExitEvent = procedure(Sender:Tobject) of object;

  TDblClickEvent = procedure(Sender:Tobject; Button:TMouseButton; Shift:TShiftState; x,y:integer) of object;

  TKeyDownEvent = procedure(Sender: Tobject; Shift: TShiftState; Key: word) of object;
  TKeyUpEvent = procedure(Sender:Tobject; Shift:TShiftState; Key:word) of object;

  TLjnWebBrowserEditor = class(TWebBrowser)
  private
    FDocument : IHTMLDocument2;
    FTheWind: IHTMLWindow2;
    FCommand : IOleCommandTarget;  //MSHTML IOLECommandTarget interface
    FEditModeOn: Boolean;
    NullVariant: OleVariant;
    FCodePage: Word;
    FAppendData: Boolean;
    FContentTransfer: string;

    FKeyDownEvent: TKeyDownEvent;
    FAlt_S_DownEvent: TKeyDownEvent;
    FCtr_Enter_DownEvent: TKeyDownEvent;
    FEnter_UpEvent: TKeyUpEvent;
    FEnter_DownEvent: TKeyDownEvent;

    FEvents: TMSHTMLHTMLDocumentEvents;

    FParentHandle: Cardinal;
    FParentForm: TForm;


    function GetEditMode: Boolean;
    function GetHtmlColorFromColor(AColor:TColor):string;
    procedure SetEditMode(const Value: Boolean);
    procedure DoNavigateComplete2(Sender: TObject; const pDisp: IDispatch; const URL: OleVariant);

    procedure DefineEvents;
    procedure OnMouseDown(Sender:TObject);
    procedure OnMouseUp(Sender:TObject);
    procedure OnMouseMove(Sender:TObject);
    procedure OnMouseOver(Sender:TObject);
    procedure OnMouseOut(Sender:TObject);
    function OnClick(Sender:TObject):WordBool;
    function OnDbClick(Sender:TObject):WordBool;
    function OnSelectStart(Sender:TObject):WordBool;
    procedure OnFocusOut(Sender:TObject);
    procedure OnFocusIn(Sender:TObject);
    function OnContextMenu(Sender:TObject):WordBool;
    function OnKeyPress(Sender:TObject):WordBool;
    procedure OnKeyDown(Sender:TObject);
    procedure OnKeyUp(Sender:TObject);
    function GetOuterHTML: WideString;
    function GetouterText: WideString;
    function GetOuterHtml2: string;
    function GetOuterHtml3: string;

    //add by pcplayer
    function InsertTargetBlank(WS: WideString):WideString;
    function GetInnerHTML: WideString;
    procedure SetInnerHTML(const Value: WideString);
    procedure SetOuterText(const Value: WideString);
    procedure SetOuterHTML(const Value: WideString);
    function GetInnerText: WideString;
    procedure SetInnerText(const Value: WideString);

    //额外的文字处理
    procedure CheckHtmlFontProperty(var WS: string);

    procedure InternalLoadDocumentFromStream(const Stream: TStream); //从 Stream 读入文档到 Browser
    procedure InternalSaveBodyHTMLToStream(const Stream: TStream);
    procedure LoadStringIntoBrowser(const HTML: string);
    function GetReadyState: WideString;
    function GetTempPath: string;
    procedure SetTempPath(const Value: string);
    function GetContentTransfer: string;
    procedure SetContentTransfer(const Value: string);
    procedure SetMyParentHandle(const Value: Cardinal);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure OpenBrowser;
    procedure SetFontSize(ASize:Integer);
    procedure SetFontColor(AColor:TColor);
    procedure SetFontBgColor(AColor:TColor);
    procedure SetFontBold;
    procedure SetFontName(FontName:string);
    procedure SetFontItalic;
    procedure SetFontUnderline;
    procedure SetFontStrikeThrough; //删除线
    procedure SetSubscript; //下标
    procedure SetSuperscript; //上标

    procedure SetJustifyCenter; //居中
    procedure SetJustifyLeft; //左对齐
    procedure SetJustifyRight; //右对齐

    procedure DoCopy; //复制
    procedure DoCut; //剪切
    procedure DoDelete; //删除
    procedure DoPaste; //粘贴
    procedure DoSelectAll; //全选

    procedure DoUnDo;
    procedure DoRedo;
    procedure DoRemoveFormat; //去除选中的字符的格式
    procedure DoRemoveParaFormat; //去除段落的格式
    procedure DoSaveAs;

    procedure DoInsertImage;
    procedure DoInsertLocalImage(AFileName: string);
    procedure DoCreateLink; //在选中的字符上插入超级链接
    procedure DoUnlink; //去掉选中的链接
    procedure DoInsertHorizontalRule; //插入横线
    procedure DoInsertTable(Rows:Integer = 4;Cols:Integer = 4; AWidth:WideString = '100%';
      Border:Integer = 1; BgColor: TColor = clWhite; Caption:WideString='';
      Padding:Integer = 0;SPacing:Integer = 0); //插入表格

    procedure SetIndent;
    procedure SetInsertParagraph; //换行！实际上是换段落
    procedure SetFocusMe;
    procedure ClearHTML;

    procedure ClearHTMLWithNoFocus;
    procedure SetBorderNone;

    procedure EnableBrowserEvent;
  published
    property LjnEditModeOn: Boolean read GetEditMode write SetEditMode;


    property LjnDocument:IHTMLDocument2 read FDocument;
    property LjnOnKeyDown: TKeyDownEvent read FKeyDownEvent write FKeyDownEvent;
    property LjnOnAlt_S_DownEvent: TKeyDownEvent read FAlt_S_DownEvent write FAlt_S_DownEvent;
    property LjnOnCtr_Enter_DownEvent: TKeyDownEvent read FCtr_Enter_DownEvent write FCtr_Enter_DownEvent;
    property LJnOnEnter_UpEvent: TKeyUpEvent read FEnter_UpEvent write FEnter_UpEvent;
    property LjnOnEnter_DownEvent: TKeyDownEvent read FEnter_DownEvent write FEnter_DownEvent;
    property LjnOuterHtml: WideString read GetOuterHTML write SetOuterHTML;
    property LjnOuterText: WideString read GetouterText write SetOuterText;
    property LjnOuterHtml2: string read GetOuterHtml2;
    property LjnOuterHtml3: string read GetOuterHtml3;
    property LjnInnerHTML: WideString read GetInnerHTML write SetInnerHTML;
    property LjnInnerText: WideString read GetInnerText write SetInnerText;
    property LjnReadyState: WideString read GetReadyState;
    property LjnTempPath: string read GetTempPath write SetTempPath;
    property LjnContentTransfer: string read GetContentTransfer write SetContentTransfer;
    property LjnParentHandle: Cardinal write SetMyParentHandle;
    property LjnParentForm: TForm read FParentForm write FParentForm;
  end;

const MyBrowserMessage = WM_USER+102;

procedure Register;

implementation

uses Clipbrd; {Clipbrd 这个单元必须加在 implementation 下，并且必须是这里的 uses 列表中的最后一个}

procedure Register;
begin
  RegisterComponents('Internet', [TLjnWebBrowserEditor]);
end;

{ TLjnWebBrowserEditor }

procedure ShowTextInWebBrowser(WS:WideString; WB:IHTMLDocument2; AppendData:Boolean);
var
  vv: Variant;
  WStr: WideString;
  i: Integer;
  FuHao: string;
begin
  if not Assigned(WB) then
  begin
    Exit;
  end;

  vv := VarArrayCreate([0, 0], varVariant);

  if AppendData then
  begin
    WStr := WB.body.outerHTML;
    i := AnsiPOS('<BODY>',WStr);
    if i>0 then
    begin
      Delete(WStr,i,6);
    end;

    i := AnsiPOS('</BODY>',WStr);
    if i > 0 then
    begin
      Delete(WStr,i,7);
    end;

    WStr := Trim(WStr);

    FuHao := Copy(WStr,Length(WStr) - 3,4);
    if FuHao = '</P>' then Delete(WStr,Length(WStr) - 3,4); //去掉结尾的 </P>，让第二条消息的换行不会离太远。

    WStr := {WStr + '<BR />' +} WS;
  end
  else
  begin
    WB.execCommand('SELECTALL', False, 0);
    WB.execCommand('Delete', False, 0);
    WStr := WS;
  end;
  vv[0] := WStr;

  WB.write(pSafearray(TVarData(vv).VArray));
  WB.parentWindow.scrollBy(0,5000);

//  HTMLDocument.close;
end;



constructor TLjnWebBrowserEditor.Create(AOwner: TComponent);
begin
  inherited;
  Self.OnNavigateComplete2 := DoNavigateComplete2;
end;

destructor TLjnWebBrowserEditor.Destroy;
begin
  if Assigned(FDocument) then FDocument.close;
  FDocument:=nil;
  FCommand:=nil;
  inherited;
end;

procedure TLjnWebBrowserEditor.DoNavigateComplete2(Sender: TObject;
  const pDisp: IDispatch; const URL: OleVariant);
begin
  //WebBrowser NavigateComplete2 事件
  FDocument := Self.Document as IHTMLDocument2;
  if Assigned(FDocument) then begin
     FTheWind := FDocument.parentWindow;

    //Self.SetEditMode(Self.FEditModeOn);
  end;
  //这句话放在这里，可以设置浏览器没边框。但是当程序设置浏览器为编辑器后，又有边框了。因此，要在设置为编辑器以后再执行这句话，才能没有边框。
  //FDocument.body.style.borderStyle := 'none';

  if FEditModeOn then
  begin
    FDocument.DesignMode := 'On';
  end;
  PostMessage(Self.FParentHandle, MyBrowserMessage, 0, 0);
  VCL.Forms.Application.ProcessMessages; //这里没加这句话，在 google 输入法底下可以输入字符。加了，纯英文键盘不能输入，google 输入法也不能输入。
end;

function TLjnWebBrowserEditor.GetContentTransfer: string;
begin
  Result := Self.FContentTransfer;
end;

function TLjnWebBrowserEditor.GetEditMode: Boolean;
begin
  Result:=FEditModeOn;
end;

function TLjnWebBrowserEditor.GetHtmlColorFromColor(
  AColor: TColor): string;
var
  AColorChar:array[0..SizeOf(TColor)*2] of AnsiChar;
  AColorStr: AnsiString;  //以前是 string，但在 Delphi 2010 底下，必须是 AnsiString
  HtmlColorLength:Integer;
begin
  FillChar(AColorChar, SizeOf(TColor) * 2 -1, 0);
  BinToHex(@AColor,AColorChar,SizeOf(TColor));

  HtmlColorLength:=6; //给 WebBrowser 的颜色命令的颜色值是 RGB 每个2位共6位
  SetLength(AColorStr,HtmlColorLength);  //
  Move(AColorChar,AColorStr[1],HtmlColorLength);

  Result:=AColorStr;
end;

procedure TLjnWebBrowserEditor.OpenBrowser;
begin
  Self.Navigate('about:blank');
end;

procedure TLjnWebBrowserEditor.SetFontBgColor(AColor: TColor);
var
  AColorStr:string;
begin
  AColorStr:=GetHtmlColorFromColor(AColor);
  if Assigned(FDocument) then
    FDocument.execCommand('BackColor',False,AColorStr);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetBorderNone;
begin
  if Assigned(FDocument) then
  begin
    FDocument.body.style.borderStyle := 'none';
  end;
end;

procedure TLjnWebBrowserEditor.SetContentTransfer(const Value: string);
begin
  Self.FContentTransfer := Value;
end;

procedure TLjnWebBrowserEditor.SetEditMode(const Value: Boolean);
var
  AMode:string;
begin
  FEditModeOn := Value;
  if Value then AMode:='On' else AMode:='Off';
  if not Assigned(FDocument) then Exit;

  FDocument.DesignMode := AMode;
  SetFocusMe;


  //FDocument.body.style.borderStyle := 'none';

  {--------------------------------------------------------------------------
    设置边框为 none  MSDN 这样说：
    When you set designMode to "on" the body element is immediately emptied (get_body returns NULL).
    You need to wait until the document's readyState is "complete" again. Once the doc is done reloading, get_body works.

    因此有以下代码，测试通过：2008-8-2

    以下代码在升级为 IE7 以后，一直无法跳出循环！暂时屏蔽掉，以后再看是什么原因。2009-2-7
  ---------------------------------------------------------------------------}
  {
  while True do
  begin
    if LowerCase(FDocument.readyState) = 'complete' then
    begin
      FDocument.body.style.borderStyle := 'None';
      Break;
    end;
    Sleep(100);
    VCL.Forms.Application.ProcessMessages; //必须加上这个，否则一直等不到 readyState 变回 complete
  end;
  }
end;

procedure TLjnWebBrowserEditor.SetFontColor(AColor: TColor);
var
  AColorStr:string;
begin
  AColorStr:=GetHtmlColorFromColor(AColor);
  if Assigned(FDocument) then
    FDocument.execCommand('ForeColor',False,AColorStr);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetFontSize(ASize: Integer);
begin
  if Assigned(FDocument) then
    FDocument.execCommand('FontSize',False,ASize);

  SetFocusMe;
  //备注：根据 MSDN FontSize 最大为 7
  //经过试验，大于7的值不会导致错误，但字体不再变大
end;

procedure TLjnWebBrowserEditor.SetFontBold;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Bold',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetFontName(FontName: string);
begin
  if Assigned(FDocument) then
    FDocument.execCommand('FontName',False,FontName);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetIndent;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Indent',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetInsertParagraph;
begin
  //强制换行，相当于硬回车。Paragraph 是 段落 的意思.
  if Assigned(FDocument) then
    FDocument.execCommand('InsertParagraph',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetFontItalic;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Italic',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetFontUnderline;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Underline',False,0);

  SetFocusMe;
end;



procedure TLjnWebBrowserEditor.SetJustifyCenter;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('JustifyCenter',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetJustifyLeft;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('JustifyLeft',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetJustifyRight;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('JustifyRight',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetMyParentHandle(const Value: Cardinal);
begin
  FParentHandle := Value;
end;

procedure TLjnWebBrowserEditor.DoCopy;
var
  spSelObj:IHTMLSelectionObject;
  spTxtRng:IHTMLTxtRange;
begin
{-------------------------------------------------------------------------------
  此命令不能被执行的原因居然是没有做 OleInitialize(nil);
  但同样没有做 OleInitialize(nil); 其它命令就能执行。真是怪哉。

---------------------------------------------------------------------------   }
  if Assigned(FDocument) then
     FDocument.execCommand('Copy' ,False,0);


  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoCut;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Cut',False,0);

  SetFocusMe;
  {ToDo: 剪切操作也不能工作}
end;
procedure TLjnWebBrowserEditor.DoDelete;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Delete',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoPaste;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Paste',False,0);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoUnDo;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Undo',NullVariant,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoRedo;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Redo',False,0);

  SetFocusMe;
end;



procedure TLjnWebBrowserEditor.SetFontStrikeThrough;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('StrikeThrough',NullVariant,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetSubscript;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Subscript',NullVariant,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetSuperscript;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Superscript',NullVariant,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetTempPath(const Value: string);
begin

end;

procedure TLjnWebBrowserEditor.DoRemoveFormat;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('RemoveFormat',NullVariant,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoRemoveParaFormat;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('RemoveParaFormat',NullVariant,NullVariant);

  SetFocusMe;
  {ToDo: 去段落格式似乎不工作}
end;

procedure TLjnWebBrowserEditor.DoSaveAs;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('SaveAs',True,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.SetFocusMe;
begin
  if FEditModeOn then
    Self.SetFocus;

  if Assigned(FDocument) then
    (FDocument.parentWindow).focus;
end;

procedure TLjnWebBrowserEditor.DoInsertImage;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('InsertImage',True,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoCreateLink;
var
  CurrentSel: IHTMLTxtRange;
  S: WideString;
  i: Integer;
begin
  if Assigned(FDocument) then
  begin
    FDocument.execCommand('CreateLink',True,NullVariant);

    {---------------------------------------------------------------------
      以下代码让超级链接变成点了后开新窗口。 add by pcplayer 2006-12-22
      以下代码测试通过，屏蔽不用。这里不对 HTML 做改动。
      需要改动
    -------------------------------------------------------------------------}
    {
    CurrentSel := FDocument.selection.createRange as IHTMLTxtRange;
    S := CurrentSel.htmlText;
    i := POS('<A', S);
    if i > 0 then
    begin
      Insert('target="_blank" ', S, i+3);  //直接在这里插入一个 target 描述后，替换掉当前的 HTML
      CurrentSel.pasteHTML(S);
    end;
    }
  end;
  

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoUnlink;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('Unlink',False,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.EnableBrowserEvent;
begin
  Self.DefineEvents;
end;

procedure TLjnWebBrowserEditor.DoInsertHorizontalRule;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('InsertHorizontalRule',False,NullVariant);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DoInsertTable(Rows:Integer = 4;Cols:Integer = 4; AWidth:WideString = '100%';
      Border:Integer = 1; BgColor: TColor = clWhite; Caption:WideString='';
      Padding:Integer = 0;SPacing:Integer = 0); //插入表格
var
  spTable:IHTMLTable;
  spElm:IHTMLElement;
  spRow:IHTMLTableRow;
  spCell:IHTMLTableCell;
  spSelObj:IHTMLSelectionObject;
  spTxtRng:IHTMLTxtRange;

  S, AColorStr:WideString;
  i,j:Integer;
begin
{---------------------------------------------------------------------
  插入表格。MSHTML 对插入表格的命令不支持，因此只好自己写。
-------------------------------------------------------------------------------}
  if not Assigned(FDocument) then
  begin
    SetFocusMe;
    Exit;
  end;

  spElm:=FDocument.createElement('Table');
  spTable:=spElm as IHtmlTable;

  for i:=0 to Rows-1 do
  begin
    spRow:=(spTable.insertRow(i) as IHTMLTableRow);
    for j:=0 to Cols-1 do
    begin
      spRow.insertCell(j);
    end;
  end;

  spTable.border:=Border;
  spTable.width:=AWidth;


  spTable.cellSpacing:=IntToStr(Padding);
  spTable.cellPadding:=IntToStr(Spacing);

  // spTable.createCaption.align:='Center';  caption 似乎没用
  //spTable.createCaption.

  if BgColor<>clWhite then
  begin
    AColorStr:=GetHtmlColorFromColor(BgColor);
    spTable.bgColor:=AColorStr;
  end;

  spSelObj:=FDocument.selection;
  spTxtRng:=spSelObj.createRange as IHTMLTxtRange;

  S:=spElm.outerHTML;
  spTxtRng.pasteHTML(S);

  //--------------释放接口----------------------------------------------
  spTable:=nil;
  spElm:=nil;
  spRow:=nil;
  spCell:=nil;
  spSelObj:=nil;
  spTxtRng:=nil;

  //-----------------------------------------------------------------------

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.DefineEvents;
begin
{---------------------------------------------------------------------
  add at 2006-3-30 by pcplayer for catch event
--------------------------------------------------------------------------}
  if not Assigned (Self.FDocument) then
    Exit;

  if FEvents<>nil then
     FEvents.Free;

  FEvents := TMSHTMLHTMLDocumentEvents.Create(Self.FParentForm);   //将这个 Self 改为它上面的 TForm 也不行。


  FEvents.Connect(IUnknown(Document));    // 用这句，连中文输入都不行。

  //FEvents.Connect(Self.FDocument);       //用这句，英文输入不行，中文可以。



  FEvents.OnMouseDown := OnMouseDown;
  FEvents.OnMouseUp := OnMouseUp;
  FEvents.OnMouseMove := OnMouseMove;
  FEvents.OnMouseOver := OnMouseOver;
  FEvents.OnMouseOut := OnMouseOut;
  FEvents.OnClick := OnClick;
  FEvents.ondblclick := OnDbClick;
  FEvents.OnSelectStart := OnSelectStart;
  FEvents.OnKeyPress := OnKeyPress;
  FEvents.OnKeyDown := OnKeyDown;
  FEvents.OnKeyUp := OnKeyUp;
  FEvents.OnContextMenu := OnContextMenu;
  FEvents.OnFocusOut := OnFocusOut;
  FEvents.OnFocusIn := OnFocusIn;

  Self.SetFocusMe;
end;

function TLjnWebBrowserEditor.OnClick(Sender: TObject): WordBool;
begin
  Result := True;
end;

function TLjnWebBrowserEditor.OnContextMenu(Sender: TObject): WordBool;
begin

end;

function TLjnWebBrowserEditor.OnDbClick(Sender: TObject): WordBool;
begin

end;

procedure TLjnWebBrowserEditor.OnFocusIn(Sender: TObject);
begin

end;

procedure TLjnWebBrowserEditor.OnFocusOut(Sender: TObject);
begin

end;

procedure TLjnWebBrowserEditor.OnKeyDown(Sender: TObject);
begin
  if Assigned(FKeyDownEvent) then
    FKeyDownEvent(self,[],FTheWind.Event.KeyCode);


  if FTheWind.Event.ctrlKey then
  begin
    if FTheWind.Event.KeyCode = VK_Return then
      if Assigned(FCtr_Enter_DownEvent) then FCtr_Enter_DownEvent(self,[ssCtrl],FTheWind.Event.KeyCode);
  end
  else
  if FTheWind.Event.AltKey then
  begin
    if (FTheWind.Event.KeyCode = Ord('s')) or (FTheWind.Event.KeyCode = Ord('S')) then
    begin
      if Assigned(FAlt_S_DownEvent) then
      begin
        FAlt_S_DownEvent(self,[ssAlt],Ord('S'));
      end;
    end;
  end
  else
  if FTheWind.event.keyCode = VK_Return then
  begin
    if Assigned(FEnter_DownEvent) then FEnter_DownEvent(self,[],FTheWind.Event.KeyCode);
  end;
end;

function TLjnWebBrowserEditor.OnKeyPress(Sender: TObject): WordBool;
begin
  Result := True; //here must return true else webbrowser can not receive key input.
end;

procedure TLjnWebBrowserEditor.OnKeyUp(Sender: TObject);
begin
//  if  FTheWind.Event.KeyCode = VK_Return then
//    if Assigned(FEnter_UpEvent) then FEnter_UpEvent(self,[],13);
end;

procedure TLjnWebBrowserEditor.OnMouseDown(Sender: TObject);
begin

end;

procedure TLjnWebBrowserEditor.OnMouseMove(Sender: TObject);
begin

end;

procedure TLjnWebBrowserEditor.OnMouseOut(Sender: TObject);
begin

end;

procedure TLjnWebBrowserEditor.OnMouseOver(Sender: TObject);
begin

end;

procedure TLjnWebBrowserEditor.OnMouseUp(Sender: TObject);
begin

end;

function TLjnWebBrowserEditor.OnSelectStart(Sender: TObject): WordBool;
begin

end;

function TLjnWebBrowserEditor.GetOuterHTML: WideString;
begin
  Result := '';
  if Assigned(FDocument) then
    Result := FDocument.body.outerHTML;
end;

function TLjnWebBrowserEditor.GetouterText: WideString;
begin
  Result := '';
  if Assigned(FDocument) then
    Result := FDocument.body.outerText;
end;

function TLjnWebBrowserEditor.GetReadyState: WideString;
begin
  Result := '';
  if Assigned(FDocument) then
  begin
    Result := FDocument.readyState;
  end;
end;

function TLjnWebBrowserEditor.GetTempPath: string;
begin
  Result := '';
end;

procedure TLjnWebBrowserEditor.ClearHTML;
var
  vv: Variant;
begin
  {
  vv := VarArrayCreate([0, 0], varVariant);
  vv[0] := '';
  FDocument.designMode:='Off';
  FDocument.write(pSafearray(TVarData(vv).VArray));
  FDocument.designMode:='On';
  }

  //changed by pcplayer 2006-5-6
  //DoSelectAll;
  if Assigned(FDocument) then
    FDocument.execCommand('SELECTALL',False,0);
  DoDelete;
end;

function TLjnWebBrowserEditor.GetOuterHtml2: string;
var
  i,j:Integer;
  WS, WSTmp,STmp, AOuterText: string;
  ACodePage:Word;
  CodePageStr:string;
  PicEncode:TLjnMessageEncodePic;
begin
{-------------------------------------------------------------------------
  聊天用编辑器，需要把浏览器自己给出来的 <PRE> 和 <P> 都干掉，避免第一条消息
  和第二条消息之间有<PRE>...</PRE> 或 <P>...</P>。
  但同一条消息中间的 <P> 这样的换行要留着。
  在这里把图片也编码到字符串里去（MIME），把 CodePage 也加进去。
----------------------------------------------------------------------------}
  Result := '';
  if not Assigned(FDocument) then Exit;

  WS := FDocument.body.outerHTML;

  AOuterText := FDocument.body.outerText;
  if (AOuterText = '') and (Pos('IMG', WS) = 0) then //先检验一下是否有输入。如果采用 outerHTML 则因为有 HTML 标记，不好判断。2008-8-20
  begin
    Self.ClearHTML;
    Self.SetFocusMe;
    Exit;
  end;

  Self.CheckHtmlFontProperty(WS);  //处理一下字符串的字体问题。这段程序是新加的，会不会导致其它问题呢？2008-8-22
  WSTmp := WS;

  WSTmp := Trim(WSTmp);
  i := POS('<P>&nbsp;</P>', WSTmp);
  while i > 0 do
  begin
    Delete(WSTmp, i, 13);
    i := POS('<P>&nbsp;</P>', WSTmp);
  end;

  STmp := Char($D) + Char($A);
  i := POS(STmp, WSTmp);
  while i > 0 do
  begin
    Delete(WSTmp, i, 2);
    i := POS(STmp, WSTmp);
  end;

  WSTmp := Trim(WSTmp);
  if WSTmp = '<BODY></BODY>' then
  begin
    WS := '';
    self.ClearHTML;
    Exit;
  end;

  if WS = '<BODY><P>&nbsp;</P></BODY>' then
  begin
    WS := '';
    Exit;
  end;

  //HLogAdd(WS);
  
  WS := Trim(WS);
  i :=Pos('<BODY>',WS);
  if i>0 then
  begin
    Delete(WS,i,6);
  end;

  //对于回车发送的，要去除最后一行的 <P>&nbsp;</P>
  i := POS('<P>&nbsp;</P></BODY>', WS);
  if i > 0 then
  begin
    Delete(WS, i, 13);
  end;
  
  i := POS('</BODY>',WS);
  if i>0 then
  begin
    Delete(WS,i,7);
  end;


  i := POS('<P>',WS);
  if i = 1 then
  begin
    Delete(WS,i,3);
  end;

  i := POS('</P>',WS);  //对于多行，把第一行的换行去掉。
  if i > 0 then
  begin
    Delete(WS,i,4);
  end;

  i := POS('<PRE>',WS);
  if i = 1 then
  begin
    Delete(WS,i,5);
  end;

  WS := StringReplace(WS,'<PRE>','<P>',[rfReplaceAll]);

  {
  i := POS('</PRE>',WS);
  while i > 0 do
  begin
    Delete(WS,i,6);
    i := POS('</PRE>',WS); //去掉所有的 </PRE>
  end;
  }

  WS := StringReplace(WS,'</PRE>','</P>',[rfReplaceAll]);


  {//去掉所有的 <Font ....></Font> 这样的没有意义的格式。add by pcplayer 2006-8-24
  i := POS('<FONT', WS);
  while i > 0 do
  begin
    j := i;

    i := PosEx('>', WS, i + 1); //找到 <FONT ...> 的结尾处
    if i > 0 then
    begin
      if PosEx('</FONT>', WS, i) = (i + 1) then //如果上一个 <FONT ..>的结尾紧跟 </FONT>
      begin
        Delete(WS, j, i + 8); // 干掉 <Font ....></Font>
      end;
    end
    else Break;

    //开始找下一组
    i := PosEx('<FONT', WS, i);
  end;
  }
  
  //现在发现切换输入文字的颜色后，结尾处可能会出现 :
  //<FONT face=3DArial><FONT size=3D2><FONT color=3D#ff0000>=BA=BA=D7=D6</FO=
  //NT></FONT></FONT>
  //<FONT face=3DArial><FONT size=3D2><FONT color=3D#ff0000>
  //这样的结尾有 <FONT> 而没有 </FONT> 的标记，要去掉。

  i := POS('<FONT',WS);
  j := i;
  while i > 0 do
  begin
    i := POSEx('</FONT>', WS, i + 1);
    if i = 0 then  //找到 <FONT> 找不到 </FONT>
    begin
      i := POSEx('>', WS, j); //找到这个 <FONT 的结尾 >
      Delete(WS, j, i - j + 1);
      i := POSEx('<FONT', WS, j);
    end
    else  //找到一个 </FONT>，则继续找下一对 <FONT>
    begin
      i := POSEx('<FONT', WS, i);
      j := i;
    end;
  end;

  WS := WS + '<P />';  //pcplayer 2006-9-26 行尾加个软回车看看

  WS := InsertTargetBlank(WS); //为超级链接插入 target="_blank"

  PicEncode := TLjnMessageEncodePic.Create;
  try
    WS := PicEncode.MessageEmbedPic(WS);
  finally
    PicEncode.Free;
  end;

  ACodePage := GetACP;
  SetLength(CodePageStr,SizeOf(Word));
  Move(ACodePage,CodePageStr[1],SizeOf(Word));

  Result := CodePageStr + WS;
end;

procedure TLjnWebBrowserEditor.DoSelectAll;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('SELECTALL',False,0);

  SetFocusMe;
end;

function TLjnWebBrowserEditor.InsertTargetBlank(
  WS: WideString): WideString;
var
  i: Integer;
  //S: WideString;
  S: string;
begin
{-------------------------------------------------------------------------
  搜索超级链接，在里面加上 Target="_Blank" 让看到的人点链接的时候弹新窗口
-----------------------------------------------------------------------------}
  S := WS;
  i := Pos('<A href', S);
  while i > 0 do
  begin
    i :=  PosEx('>', S, i);
    Insert(' target="_blank" ', S, i);
    i := PosEx('<A href', S, i);
  end;

  Result := S;
end;

procedure TLjnWebBrowserEditor.InternalLoadDocumentFromStream(
  const Stream: TStream);
var
  PersistStreamInit: IPersistStreamInit;
  StreamAdapter: IStream;
begin
  Assert(Assigned(Self.Document));
  // Get IPersistStreamInit interface on document object
  if Self.Document.QueryInterface(
    IPersistStreamInit, PersistStreamInit
  ) = S_OK then
  begin
    // Clear document
    if PersistStreamInit.InitNew = S_OK then
    begin
      // Get IStream interface on stream
      StreamAdapter:= TStreamAdapter.Create(Stream);
      // Load data from Stream into WebBrowser
      PersistStreamInit.Load(StreamAdapter);
    end;
  end;
end;

procedure TLjnWebBrowserEditor.InternalSaveBodyHTMLToStream(
  const Stream: TStream);
var
  {$IFDEF UNICODE}
  Bytes: TBytes;
  {$ENDIF}
  HTMLStr: string;
  Doc: IHTMLDocument2;
  BodyElement: IHTMLElement;
begin
  Assert(Assigned(Self.Document));
  if Self.Document.QueryInterface(IHTMLDocument2, Doc) = S_OK then
  begin
    BodyElement := Doc.body;
    if Assigned(BodyElement) then
    begin
      HTMLStr := BodyElement.innerHTML;
      {$IFDEF UNICODE}
      Bytes := BytesOf(HTMLStr);
      Stream.Write(Bytes[0], Length(Bytes));
      {$ELSE}
      Stream.WriteBuffer(HTMLStr[1], Length(HTMLStr));
      {$ENDIF}
    end;
  end;
end;

procedure TLjnWebBrowserEditor.LoadStringIntoBrowser(const HTML: string);
var
  StringStream: TStringStream;
begin
{---------------------------------------------------------------------------
  加载文档到浏览器里面，从 ShowTextInWebBrowser 方法改为加载流的方法。
-----------------------------------------------------------------------------}
  StringStream := TStringStream.Create(HTML);
  try
    InternalLoadDocumentFromStream(StringStream);
  finally
    StringStream.Free;
  end;
end;

procedure TLjnWebBrowserEditor.DoInsertLocalImage(AFileName: string);
begin
{---------------------------------------------------------------------------
  没有文件名选择的提示对话框，直接把一个图片文件插入浏览器编辑器中
  execCommand('InsertImage',True,AFileName); 则是显示选择图片文件对话框。False 则不显示。
-------------------------------------------------------------------------------}
  if Assigned(FDocument) then
    FDocument.execCommand('InsertImage',False,AFileName);

  SetFocusMe;
end;

procedure TLjnWebBrowserEditor.ClearHTMLWithNoFocus;
begin
  if Assigned(FDocument) then
    FDocument.execCommand('SELECTALL',False,0);

  if Assigned(FDocument) then
    FDocument.execCommand('Delete',False,0);
end;

function TLjnWebBrowserEditor.GetInnerHTML: WideString;
begin
  if Assigned(FDocument) then Result := FDocument.body.innerHTML;
end;

procedure TLjnWebBrowserEditor.SetInnerHTML(const Value: WideString);
begin
  if Assigned(FDocument) then
    FDocument.body.innerHTML := Value;
end;

procedure TLjnWebBrowserEditor.SetOuterText(const Value: WideString);
begin
  if Assigned(FDocument) then FDocument.body.outerText := Value;
end;

procedure TLjnWebBrowserEditor.SetOuterHTML(const Value: WideString);
begin
  //if Assigned(FDocument) then FDocument.body.outerHTML := Value;  执行这句话会出 AV 错误！改造为以下代码看看。
  if Assigned(FDocument) then
  begin
    //ShowTextInWebBrowser(Value, FDocument, False);
    Self.LoadStringIntoBrowser(Value);
  end;
end;

function TLjnWebBrowserEditor.GetInnerText: WideString;
begin
  if Assigned(FDocument) then Result := FDocument.body.innerText;
end;

procedure TLjnWebBrowserEditor.SetInnerText(const Value: WideString);
begin
  if Assigned(FDocument) then FDocument.body.innerText := Value;
end;

procedure TLjnWebBrowserEditor.CheckHtmlFontProperty(
  var WS: string);
var
  i, j, k, L, M: Integer;
  S, S2: string;
begin
{------------------------------------------------------------------------------
  因为 IE 本身的 BUG 会导致有时候拿到的 OutHtml 字符串格式如：

  #$D#$A'<BODY style="BORDER-TOP-STYLE: none; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; BORDER-BOTTOM-STYLE: none"><P><FONT style="BACKGROUND-COLOR: #ffffff" face=Arial size=2></FONT>关掉这个会议。</P>'#$D#$A'<P>&nbsp;</P></BODY>'

  即 <Font ></Font> 中间没有字符，而字符串两边没有 Font 标记，导致字符的格式不对。
  本方法就是检查是否有这种情况，如果有，则修正，把字符串取出来放到 <font> 标记中间。

  Delphi 7 的 Pos 和 PosEx 处理不好 WideString 所以这里必须是 string

  2008-8-22
---------------------------------------------------------------------------------}
  S := UpperCase(WS);

  i := Pos('<FONT', S);
  j := PosEx('>', S, i + 1); //找到了 <Font> 的结尾, 然后要看后面是否紧随 </Font>
  k := PosEx('</FONT', S, j + 1);
  if (k - j) > 1 then Exit;     //<font>和</Font> 标记中间有其它字符串，说明这是正常的字符串，不用修正了。

  L := PosEx('>', S, K + 1); //找到 </FONT> 的后括号 > ，然后找下一个 < 括号。之间的字符串就是要的字符串

  M := PosEx('<', S, L + 1); //找到 < 括号

  S2 := Copy(WS, L + 1, M - L - 1); //把 </Font> 和 < 括号之间的字符串取出来。
  if S2 = '' then Exit; //如果没有，则直接退出。

  Delete(WS, L + 1, M - L - 1); //把这段字符串去掉；
  Insert(S2, WS, j + 1);  //把这段字符串插入到两个 font 标记中间。
end;

function TLjnWebBrowserEditor.GetOuterHtml3: string;
var
  PlanText, WS: WideString;
  PicEncode: TLjnMessageEncodePic;
begin
  Result := '';
  PlanText := FDocument.body.outerText;
  WS := FDocument.body.outerHTML;

  if (PlanText = '') and (AnsiPos('IMG', WS) = 0) then Exit;

  PicEncode := TLjnMessageEncodePic.Create;
  try
    PicEncode.ContentTransferEncoding := Self.FContentTransfer;
    Result := PicEncode.MessageEmbedPic(WS);
  finally
    PicEncode.Free;
  end;
end;

initialization
  Set8087CW($027F); //浏览器拉动滚动条的时候会出现异常。Floating point division by zero ，加上这个，异常消失。
  CoInitialize(nil);
  //OleInitialize(nil);

finalization
  CoUninitialize;

  try
//    OleUninitialize;
  except
  end;

end.
