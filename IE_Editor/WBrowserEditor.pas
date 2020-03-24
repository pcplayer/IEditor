unit WBrowserEditor;
{-----------------------------------------------------------------------------
  ����һ���� Ie ���������Ϊ���ı��༭���� Html �༭���Ŀؼ�

  �� IE ����������༭���棺
  1. ��һ�� TWebBrowser ������
  2. ������������ִ�� WebBrowser1.Navigate('about:blank');
  3. �� TWebBrowser ���¼���WebBrowser1NavigateComplete2 �ȡ�� IHTMLDocument2 �ӿ�
  4. �����������Ϊ�༭ģʽ������ȡ�õ� IHTMLDocument2 �ӿڵ� DesignMode ����Ϊ 'On' �ַ���
  5. Ҫ����ѡ�е��ַ��������塢��ɫ�ȣ����� IHTMLDocument2 �ӿڵ� execCommand ������
  6. execCommand ������Ҫ�ı�ѡ�е������С���򷽷��������ַ��� 'FontSize'�����Ҫ��Ϊ���壬���������� 'Bold'
  7. �����ʱ���������һ��Ҫ SetFocus ���������� Enter �س�����ֵ��������ȥ���������ʱ�򣬲����� SetFocus �ġ�
  8. ��Ϊ�༭�����Բ����ͼ��Ĵ�����Ϊ����ģʽ��
     A. ��ͼ��������� Base64 �����ֱ�Ӽ������������ﴫ����Է�
     B. ֻ��������ͼ������������ַ����͸��Է���

  9. MSHTML �Ա��֧�֡���Ҫ�Լ�д��
  
     ��ϸ�ķ������Ͳ������ο���
     http://msdn.microsoft.com/workshop/browser/mshtml/reference/ifaces/document2/execcommand.asp
     http://msdn.microsoft.com/workshop/author/dhtml/reference/commandids.asp
     http://msdn.microsoft.com/workshop/browser/mshtml/reference/ifaces/document2/document2.asp
     http://msdn.microsoft.com/workshop/browser/editing/mshtmleditor.asp

     �����и����������������һ�£�
     http://www.itwriting.com/htmleditor/index.php

  pcplayer. 2006-2-15

  ��������Ϊ֧�� delphi 2010 �Ŀؼ���pcplayer 2010-5-31
------------------------------------------------------------------------------}

interface

uses
  Winapi.Windows,  Winapi.Messages, Winapi.ActiveX,
  System.SysUtils, System.Variants, System.Classes, System.StrUtils, System.TypInfo,
  Vcl.Controls, Vcl.Consts, Vcl.Graphics, Vcl.Forms,
  MSHTML, MSHTMLEvents, Vcl.OleCtrls, SHDocVw,

  MiMeEncodePic{����ͼƬ��}{, LjnMailMessage{������};

type
  //�¼�
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

    //��������ִ���
    procedure CheckHtmlFontProperty(var WS: string);

    procedure InternalLoadDocumentFromStream(const Stream: TStream); //�� Stream �����ĵ��� Browser
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
    procedure SetFontStrikeThrough; //ɾ����
    procedure SetSubscript; //�±�
    procedure SetSuperscript; //�ϱ�

    procedure SetJustifyCenter; //����
    procedure SetJustifyLeft; //�����
    procedure SetJustifyRight; //�Ҷ���

    procedure DoCopy; //����
    procedure DoCut; //����
    procedure DoDelete; //ɾ��
    procedure DoPaste; //ճ��
    procedure DoSelectAll; //ȫѡ

    procedure DoUnDo;
    procedure DoRedo;
    procedure DoRemoveFormat; //ȥ��ѡ�е��ַ��ĸ�ʽ
    procedure DoRemoveParaFormat; //ȥ������ĸ�ʽ
    procedure DoSaveAs;

    procedure DoInsertImage;
    procedure DoInsertLocalImage(AFileName: string);
    procedure DoCreateLink; //��ѡ�е��ַ��ϲ��볬������
    procedure DoUnlink; //ȥ��ѡ�е�����
    procedure DoInsertHorizontalRule; //�������
    procedure DoInsertTable(Rows:Integer = 4;Cols:Integer = 4; AWidth:WideString = '100%';
      Border:Integer = 1; BgColor: TColor = clWhite; Caption:WideString='';
      Padding:Integer = 0;SPacing:Integer = 0); //������

    procedure SetIndent;
    procedure SetInsertParagraph; //���У�ʵ�����ǻ�����
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

uses Clipbrd; {Clipbrd �����Ԫ������� implementation �£����ұ���������� uses �б��е����һ��}

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
    if FuHao = '</P>' then Delete(WStr,Length(WStr) - 3,4); //ȥ����β�� </P>���õڶ�����Ϣ�Ļ��в�����̫Զ��

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
  //WebBrowser NavigateComplete2 �¼�
  FDocument := Self.Document as IHTMLDocument2;
  if Assigned(FDocument) then begin
     FTheWind := FDocument.parentWindow;

    //Self.SetEditMode(Self.FEditModeOn);
  end;
  //��仰��������������������û�߿򡣵��ǵ��������������Ϊ�༭�������б߿��ˡ���ˣ�Ҫ������Ϊ�༭���Ժ���ִ����仰������û�б߿�
  //FDocument.body.style.borderStyle := 'none';

  if FEditModeOn then
  begin
    FDocument.DesignMode := 'On';
  end;
  PostMessage(Self.FParentHandle, MyBrowserMessage, 0, 0);
  VCL.Forms.Application.ProcessMessages; //����û����仰���� google ���뷨���¿��������ַ������ˣ���Ӣ�ļ��̲������룬google ���뷨Ҳ�������롣
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
  AColorStr: AnsiString;  //��ǰ�� string������ Delphi 2010 ���£������� AnsiString
  HtmlColorLength:Integer;
begin
  FillChar(AColorChar, SizeOf(TColor) * 2 -1, 0);
  BinToHex(@AColor,AColorChar,SizeOf(TColor));

  HtmlColorLength:=6; //�� WebBrowser ����ɫ�������ɫֵ�� RGB ÿ��2λ��6λ
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
    ���ñ߿�Ϊ none  MSDN ����˵��
    When you set designMode to "on" the body element is immediately emptied (get_body returns NULL).
    You need to wait until the document's readyState is "complete" again. Once the doc is done reloading, get_body works.

    ��������´��룬����ͨ����2008-8-2

    ���´���������Ϊ IE7 �Ժ�һֱ�޷�����ѭ������ʱ���ε����Ժ��ٿ���ʲôԭ��2009-2-7
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
    VCL.Forms.Application.ProcessMessages; //����������������һֱ�Ȳ��� readyState ��� complete
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
  //��ע������ MSDN FontSize ���Ϊ 7
  //�������飬����7��ֵ���ᵼ�´��󣬵����岻�ٱ��
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
  //ǿ�ƻ��У��൱��Ӳ�س���Paragraph �� ���� ����˼.
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
  ������ܱ�ִ�е�ԭ���Ȼ��û���� OleInitialize(nil);
  ��ͬ��û���� OleInitialize(nil); �����������ִ�С����ǹ��ա�

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
  {ToDo: ���в���Ҳ���ܹ���}
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
  {ToDo: ȥ�����ʽ�ƺ�������}
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
      ���´����ó������ӱ�ɵ��˺��´��ڡ� add by pcplayer 2006-12-22
      ���´������ͨ�������β��á����ﲻ�� HTML ���Ķ���
      ��Ҫ�Ķ�
    -------------------------------------------------------------------------}
    {
    CurrentSel := FDocument.selection.createRange as IHTMLTxtRange;
    S := CurrentSel.htmlText;
    i := POS('<A', S);
    if i > 0 then
    begin
      Insert('target="_blank" ', S, i+3);  //ֱ�����������һ�� target �������滻����ǰ�� HTML
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
      Padding:Integer = 0;SPacing:Integer = 0); //������
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
  ������MSHTML �Բ���������֧�֣����ֻ���Լ�д��
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

  // spTable.createCaption.align:='Center';  caption �ƺ�û��
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

  //--------------�ͷŽӿ�----------------------------------------------
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

  FEvents := TMSHTMLHTMLDocumentEvents.Create(Self.FParentForm);   //����� Self ��Ϊ������� TForm Ҳ���С�


  FEvents.Connect(IUnknown(Document));    // ����䣬���������붼���С�

  //FEvents.Connect(Self.FDocument);       //����䣬Ӣ�����벻�У����Ŀ��ԡ�



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
  �����ñ༭������Ҫ��������Լ��������� <PRE> �� <P> ���ɵ��������һ����Ϣ
  �͵ڶ�����Ϣ֮����<PRE>...</PRE> �� <P>...</P>��
  ��ͬһ����Ϣ�м�� <P> �����Ļ���Ҫ���š�
  �������ͼƬҲ���뵽�ַ�����ȥ��MIME������ CodePage Ҳ�ӽ�ȥ��
----------------------------------------------------------------------------}
  Result := '';
  if not Assigned(FDocument) then Exit;

  WS := FDocument.body.outerHTML;

  AOuterText := FDocument.body.outerText;
  if (AOuterText = '') and (Pos('IMG', WS) = 0) then //�ȼ���һ���Ƿ������롣������� outerHTML ����Ϊ�� HTML ��ǣ������жϡ�2008-8-20
  begin
    Self.ClearHTML;
    Self.SetFocusMe;
    Exit;
  end;

  Self.CheckHtmlFontProperty(WS);  //����һ���ַ������������⡣��γ������¼ӵģ��᲻�ᵼ�����������أ�2008-8-22
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

  //���ڻس����͵ģ�Ҫȥ�����һ�е� <P>&nbsp;</P>
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

  i := POS('</P>',WS);  //���ڶ��У��ѵ�һ�еĻ���ȥ����
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
    i := POS('</PRE>',WS); //ȥ�����е� </PRE>
  end;
  }

  WS := StringReplace(WS,'</PRE>','</P>',[rfReplaceAll]);


  {//ȥ�����е� <Font ....></Font> ������û������ĸ�ʽ��add by pcplayer 2006-8-24
  i := POS('<FONT', WS);
  while i > 0 do
  begin
    j := i;

    i := PosEx('>', WS, i + 1); //�ҵ� <FONT ...> �Ľ�β��
    if i > 0 then
    begin
      if PosEx('</FONT>', WS, i) = (i + 1) then //�����һ�� <FONT ..>�Ľ�β���� </FONT>
      begin
        Delete(WS, j, i + 8); // �ɵ� <Font ....></Font>
      end;
    end
    else Break;

    //��ʼ����һ��
    i := PosEx('<FONT', WS, i);
  end;
  }
  
  //���ڷ����л��������ֵ���ɫ�󣬽�β�����ܻ���� :
  //<FONT face=3DArial><FONT size=3D2><FONT color=3D#ff0000>=BA=BA=D7=D6</FO=
  //NT></FONT></FONT>
  //<FONT face=3DArial><FONT size=3D2><FONT color=3D#ff0000>
  //�����Ľ�β�� <FONT> ��û�� </FONT> �ı�ǣ�Ҫȥ����

  i := POS('<FONT',WS);
  j := i;
  while i > 0 do
  begin
    i := POSEx('</FONT>', WS, i + 1);
    if i = 0 then  //�ҵ� <FONT> �Ҳ��� </FONT>
    begin
      i := POSEx('>', WS, j); //�ҵ���� <FONT �Ľ�β >
      Delete(WS, j, i - j + 1);
      i := POSEx('<FONT', WS, j);
    end
    else  //�ҵ�һ�� </FONT>�����������һ�� <FONT>
    begin
      i := POSEx('<FONT', WS, i);
      j := i;
    end;
  end;

  WS := WS + '<P />';  //pcplayer 2006-9-26 ��β�Ӹ���س�����

  WS := InsertTargetBlank(WS); //Ϊ�������Ӳ��� target="_blank"

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
  �����������ӣ���������� Target="_Blank" �ÿ������˵����ӵ�ʱ���´���
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
  �����ĵ�����������棬�� ShowTextInWebBrowser ������Ϊ�������ķ�����
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
  û���ļ���ѡ�����ʾ�Ի���ֱ�Ӱ�һ��ͼƬ�ļ�����������༭����
  execCommand('InsertImage',True,AFileName); ������ʾѡ��ͼƬ�ļ��Ի���False ����ʾ��
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
  //if Assigned(FDocument) then FDocument.body.outerHTML := Value;  ִ����仰��� AV ���󣡸���Ϊ���´��뿴����
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
  ��Ϊ IE ����� BUG �ᵼ����ʱ���õ��� OutHtml �ַ�����ʽ�磺

  #$D#$A'<BODY style="BORDER-TOP-STYLE: none; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; BORDER-BOTTOM-STYLE: none"><P><FONT style="BACKGROUND-COLOR: #ffffff" face=Arial size=2></FONT>�ص�������顣</P>'#$D#$A'<P>&nbsp;</P></BODY>'

  �� <Font ></Font> �м�û���ַ������ַ�������û�� Font ��ǣ������ַ��ĸ�ʽ���ԡ�
  ���������Ǽ���Ƿ����������������У������������ַ���ȡ�����ŵ� <font> ����м䡣

  Delphi 7 �� Pos �� PosEx ������ WideString ������������� string

  2008-8-22
---------------------------------------------------------------------------------}
  S := UpperCase(WS);

  i := Pos('<FONT', S);
  j := PosEx('>', S, i + 1); //�ҵ��� <Font> �Ľ�β, Ȼ��Ҫ�������Ƿ���� </Font>
  k := PosEx('</FONT', S, j + 1);
  if (k - j) > 1 then Exit;     //<font>��</Font> ����м��������ַ�����˵�������������ַ��������������ˡ�

  L := PosEx('>', S, K + 1); //�ҵ� </FONT> �ĺ����� > ��Ȼ������һ�� < ���š�֮����ַ�������Ҫ���ַ���

  M := PosEx('<', S, L + 1); //�ҵ� < ����

  S2 := Copy(WS, L + 1, M - L - 1); //�� </Font> �� < ����֮����ַ���ȡ������
  if S2 = '' then Exit; //���û�У���ֱ���˳���

  Delete(WS, L + 1, M - L - 1); //������ַ���ȥ����
  Insert(S2, WS, j + 1);  //������ַ������뵽���� font ����м䡣
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
  Set8087CW($027F); //�����������������ʱ�������쳣��Floating point division by zero ������������쳣��ʧ��
  CoInitialize(nil);
  //OleInitialize(nil);

finalization
  CoUninitialize;

  try
//    OleUninitialize;
  except
  end;

end.
