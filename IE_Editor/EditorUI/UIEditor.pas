unit UIEditor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  MSHTML,
  Dialogs, ExtCtrls, ImgList, {JvExControls, JvComponent,
  JvButton, JvTransparentButton,} RzButton, StdCtrls, RzCmboBx, Mask, RzEdit,
  RzSpnEdt, ActnList, OleCtrls, SHDocVw, WBrowserEditor, RzPanel,
  pngimage, pngextra, VnGlobal, StdActns, Menus, RzBckgnd, NLUser;

type
  TFmIEditor = class(TForm)
    Panel1: TPanel;
    ImageList1: TImageList;
    ColorDialog1: TColorDialog;
    ActionList1: TActionList;
    AcOpen: TAction;
    AcEdit: TAction;
    AcBrowse: TAction;
    AcSave: TAction;
    AcCopy: TAction;
    AcPaste: TAction;
    AcCut: TAction;
    AcDelete: TAction;
    AcLeft: TAction;
    AcCenter: TAction;
    AcRight: TAction;
    AcHuanHang: TAction;
    AcLink: TAction;
    AcLinkDrop: TAction;
    AcInsertPic: TAction;
    AcTable: TAction;
    AcLine: TAction;
    AcFontColor: TAction;
    AcFontBgColor: TAction;
    AcItalic: TAction;
    AcUnderLine: TAction;
    AcBold: TAction;
    AcUpper: TAction;
    AcLower: TAction;
    Panel2: TPanel;
    Panel3: TPanel;
    Image6: TImage;
    Image7: TImage;
    Image9: TImage;
    Image10: TImage;
    Panel4: TPanel;
    LjnWebBrowserEditor1: TLjnWebBrowserEditor;
    PNGButton1: TPNGButton;
    EditCopy1: TEditCopy;
    EditCopy2: TEditCopy;
    EditCopy3: TEditCopy;
    EditCut1: TEditCut;
    FileOpen1: TFileOpen;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    CtrEntertoSend1: TMenuItem;
    RzBackground3: TRzBackground;
    RzBackground4: TRzBackground;
    RzBackground5: TRzBackground;
    RzBackground6: TRzBackground;
    RzPanel1: TRzPanel;
    FontComboBox1: TRzFontComboBox;
    SpinEdit1: TRzSpinEdit;
    RzPanel2: TRzPanel;
    AcCatchScreen: TAction;
    AcCatchScreen2: TAction;
    PopupMenu2: TPopupMenu;
    N2: TMenuItem;
    N3: TMenuItem;
    RzBackground1: TRzBackground;
    RzToolButton1: TRzToolButton;
    RzToolButton2: TRzToolButton;
    RzToolButton8: TRzToolButton;
    RzToolButton9: TRzToolButton;
    RzToolButton10: TRzToolButton;
    RzToolButton13: TRzToolButton;
    RzToolButton14: TRzToolButton;
    RzToolButton15: TRzToolButton;
    RzToolButton16: TRzToolButton;
    RzToolButton17: TRzToolButton;
    RzToolButton18: TRzToolButton;
    RzToolButton19: TRzToolButton;
    RzToolButton20: TRzToolButton;
    RzToolButton21: TRzToolButton;
    RzToolButton22: TRzToolButton;
    RzToolButton3: TRzToolButton;
    RzToolButton4: TRzToolButton;
    RzToolButton5: TRzToolButton;
    RzToolButton6: TRzToolButton;
    RzToolButton7: TRzToolButton;
    procedure BtnFontColorClick(Sender: TObject);
    procedure AcOpenExecute(Sender: TObject);
    procedure AcEditExecute(Sender: TObject);
    procedure AcBrowseExecute(Sender: TObject);
    procedure AcSaveExecute(Sender: TObject);
    procedure AcLeftExecute(Sender: TObject);
    procedure AcCenterExecute(Sender: TObject);
    procedure AcRightExecute(Sender: TObject);
    procedure AcHuanHangExecute(Sender: TObject);
    procedure AcLinkExecute(Sender: TObject);
    procedure AcLinkDropExecute(Sender: TObject);
    procedure AcInsertPicExecute(Sender: TObject);
    procedure AcTableExecute(Sender: TObject);
    procedure AcLineExecute(Sender: TObject);
    procedure AcCopyExecute(Sender: TObject);
    procedure AcPasteExecute(Sender: TObject);
    procedure AcCutExecute(Sender: TObject);
    procedure AcDeleteExecute(Sender: TObject);
    procedure FontComboBox1Change(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure AcFontColorExecute(Sender: TObject);
    procedure AcFontBgColorExecute(Sender: TObject);
    procedure AcItalicExecute(Sender: TObject);
    procedure AcUnderLineExecute(Sender: TObject);
    procedure AcBoldExecute(Sender: TObject);
    procedure AcUpperExecute(Sender: TObject);
    procedure AcLowerExecute(Sender: TObject);
    procedure LjnWebBrowserEditor1DocumentComplete(Sender: TObject;
      const pDisp: IDispatch; var URL: OleVariant);
    procedure FormCreate(Sender: TObject);
    procedure PNGButton1Click(Sender: TObject);
    procedure LjnWebBrowserEditor1CommandStateChange(Sender: TObject;
      Command: Integer; Enable: WordBool);
    procedure N1Click(Sender: TObject);
    procedure CtrEntertoSend1Click(Sender: TObject);

    procedure FaceSelectMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FaceSelectMouseLeave(Sender: TObject);
    procedure EnterMenuMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure PNGButton28Click(Sender: TObject);
    procedure AcCatchScreenExecute(Sender: TObject);
    procedure AcCatchScreen2Execute(Sender: TObject);
    procedure PNGButton29MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure RzBackground1DblClick(Sender: TObject);
  private
    { Private declarations }
    FFirst: Boolean;
    FFontColor: TColor;
    FFontBgColor: TColor;
    FFontName: string;
    FFontSize: Integer;
    FFontItalic: Boolean;
    FFontBold: Boolean;
    FFontUnderLine: Boolean;
    FEnterSend:Boolean; //是否在按回车的时候发送 add by pcplayer 2006-7-14

    FOnAltS: TNotifyEvent;
    FOnCtrEnter: TNotifyEvent;
    FOnEnterUp: TNotifyEvent;
    FOnSendClick: TNotifyEvent;

    FDocument:IHTMLDocument2;

    //事件
    FOnMinMeetingForm: TNotifyEvent; //要求会议窗口最小化的事件
    FOnRestoreMeetingForm: TNotifyEvent;

    procedure EnterEditMode;
    procedure SetFontName(FontName: string);
    procedure SetFontSize(Size: Integer);

    procedure DoOnKeyDown(Sender:Tobject;Shift:TShiftState;Key:word);
    procedure DoAltS(Sender:Tobject;Shift:TShiftState;Key:word);
    procedure DoCtrEnter(Sender:Tobject;Shift:TShiftState;Key:word);
    procedure DoEnterUp(Sender:Tobject;Shift:TShiftState;Key:word);
    procedure DoEnterDown(Sender:Tobject;Shift:TShiftState;Key:word);

    procedure InitFont; //初始化字体
    procedure SetFontStyle;
    procedure DoInsertFace(AFileName: string);
    procedure ShowToolPanel(AStatus: Boolean);
    procedure DoAfterCatchScreen(AFileName: string);
  public
    { Public declarations }
    procedure ClearHTML;
    function GetHTML:WideString;
    function GetHTMLUtf8:string;
    function GetPageCodeStr: string;
    function GetHtml2:string;
    procedure WBSetFocus;
    procedure SetLang;
    procedure OpenBrowser;

    property OnAltS: TNotifyEvent read FOnAltS write FOnAltS;
    property OnCtrEnter: TNotifyEvent read FOnCtrEnter write FOnCtrEnter;
    property OnEnterUp:TNotifyEvent read FOnEnterUp write FOnEnterUp;
    property OnSendClick: TNotifyEvent read FOnSendClick write FOnSendClick;

    property EnterToSend:Boolean read FEnterSend;
    property NLOnMinMeetingForm: TNotifyEvent read FOnMinMeetingForm write FOnMinMeetingForm;
    property NLOnRestoreMeetingForm: TNotifyEvent read FOnRestoreMeetingForm write FOnRestoreMeetingForm;
  end;

//var
  //FmIEditor: TFmIEditor;

implementation

uses UTableProperty, UFmSelectFace, UFmFullScreen;

{$R *.dfm}

procedure TFmIEditor.BtnFontColorClick(Sender: TObject);
begin
  ColorDialog1.Execute;
end;

procedure TFmIEditor.DoAltS(Sender: Tobject; Shift: TShiftState;
  Key: word);
begin
  //ShowMessage('Alt_S!');
  if Assigned(FOnAltS) then FOnAltS(self);
  WBSetfocus;
end;

procedure TFmIEditor.DoCtrEnter(Sender: Tobject; Shift: TShiftState;
  Key: word);
begin
  //Ctr_Enter
  if Assigned(FOnCtrEnter) then FOnCtrEnter(Self);
  WBSetfocus;
end;

procedure TFmIEditor.DoOnKeyDown(Sender: Tobject; Shift: TShiftState;
  Key: word);
begin
 //OnKeyDown
end;

procedure TFmIEditor.WBSetFocus;
begin
  LjnWebBrowserEditor1.SetFocus;
  LjnWebBrowserEditor1.SetFocusMe;
end;

procedure TFmIEditor.AcOpenExecute(Sender: TObject);
begin
  OpenBrowser;
end;

procedure TFmIEditor.EnterEditMode;
begin
  LjnWebBrowserEditor1.EditModeOn:=True;
  FDocument:=LjnWebBrowserEditor1.LjnDocument;
  WBSetFocus;
end;

procedure TFmIEditor.AcEditExecute(Sender: TObject);
begin
  FFirst := True;
  EnterEditMode;
end;

procedure TFmIEditor.AcBrowseExecute(Sender: TObject);
begin
  LjnWebBrowserEditor1.EditModeOn := False;
end;

procedure TFmIEditor.AcSaveExecute(Sender: TObject);
begin
  //保存当前编辑的文本到文件，暂时不写实现
  //LjnWebBrowserEditor1.
end;

procedure TFmIEditor.AcLeftExecute(Sender: TObject);
begin
  //左对齐
  LjnWebBrowserEditor1.SetJustifyLeft;
  WBSetFocus;
end;

procedure TFmIEditor.AcCenterExecute(Sender: TObject);
begin
  //居中
  LjnWebBrowserEditor1.SetJustifyCenter;
  WBSetFocus;
end;

procedure TFmIEditor.AcRightExecute(Sender: TObject);
begin
  //右对齐
  LjnWebBrowserEditor1.SetJustifyRight;
  WBSetFocus;
end;

procedure TFmIEditor.AcHuanHangExecute(Sender: TObject);
begin
  //换段落
  LjnWebBrowserEditor1.SetInsertParagraph;
  WBSetFocus;
end;

procedure TFmIEditor.AcLinkExecute(Sender: TObject);
begin
  //插入链接
  LjnWebBrowserEditor1.DoCreateLink;
end;

procedure TFmIEditor.AcLinkDropExecute(Sender: TObject);
begin
  //去掉链接
  LjnWebBrowserEditor1.DoUnlink;
end;

procedure TFmIEditor.AcInsertPicExecute(Sender: TObject);
begin
  //插入图片
  LjnWebBrowserEditor1.DoInsertImage;
end;

procedure TFmIEditor.AcTableExecute(Sender: TObject);
var
  Rows,Cols,ABorder,Padding,SPacing,AWidth: Integer;
  AWidthStr,ACaption: WideString;
  BgColor: TColor;
begin
  //插入表格，应该弹出个对话框，设置表格的行列数字    ToDo
  //LjnWebBrowserEditor1.DoInsertTable(5,6,'60%',2,clWhite,'公司名称',1,1);
  if not Assigned(FmTableProperty) then Application.CreateForm(TFmTableProperty,FmTableProperty);

  FmTableProperty.ShowModal;

  if FmTableProperty.ModalResult = mrOK then
  begin
    FmTableProperty.GetParams(Rows,Cols,AWidth,ABorder,Padding,Spacing,BgColor,ACaption);
    AWidthStr := IntToStr(AWidth)+'%';
    LjnWebBrowserEditor1.DoInsertTable(Rows,Cols,AWidthStr,ABorder,BgColor,ACaption,Padding,Spacing);
  end;
end;

procedure TFmIEditor.AcLineExecute(Sender: TObject);
begin
  //插入横线
  LjnWebBrowserEditor1.DoInsertHorizontalRule;
end;

procedure TFmIEditor.AcCopyExecute(Sender: TObject);
begin
  //copy
  LjnWebBrowserEditor1.DoCopy;
  WBSetfocus;
end;

procedure TFmIEditor.AcPasteExecute(Sender: TObject);
begin
  //粘贴
  LjnWebBrowserEditor1.DoPaste;
  WBSetfocus;
end;

procedure TFmIEditor.AcCutExecute(Sender: TObject);
begin
  //Cut
  LjnWebBrowserEditor1.DoCut;
  WBSetfocus;
end;

procedure TFmIEditor.AcDeleteExecute(Sender: TObject);
begin
  //删除
  LjnWebBrowserEditor1.DoDelete;
  WBSetfocus;
end;

procedure TFmIEditor.SetFontName(FontName: string);
begin
  FFontName := FontName;
  LjnWebBrowserEditor1.SetFontName(FontName);
  WBSetFocus;
end;

procedure TFmIEditor.FontComboBox1Change(Sender: TObject);
begin
  SetFontName(FontComboBox1.FontName);
end;

procedure TFmIEditor.SetFontSize(Size: Integer);
begin
  FFontSize := Size;
  LjnWebBrowserEditor1.SetFontSize(Size);
  WBSetFocus;
end;

procedure TFmIEditor.SpinEdit1Change(Sender: TObject);
var
  i: Integer;
begin
  i := SpinEdit1.IntValue;
  SetFontSize(i);
end;

procedure TFmIEditor.AcFontColorExecute(Sender: TObject);
begin
  if ColorDialog1.Execute then
  begin
    FFontColor := ColorDialog1.Color;
    LjnWebBrowserEditor1.SetFontColor(ColorDialog1.Color);
    WBSetFocus;
  end;
end;

procedure TFmIEditor.AcFontBgColorExecute(Sender: TObject);
begin
  if ColorDialog1.Execute then
  begin
    FFontBgColor := ColorDialog1.Color;
    LjnWebBrowserEditor1.SetFontBgColor(FFontBgColor);
    WBSetFocus;
  end;
end;

procedure TFmIEditor.AcItalicExecute(Sender: TObject);
begin
  FFontItalic := not FFontItalic;
  LjnWebBrowserEditor1.SetFontItalic;
  WBSetFocus;
end;

procedure TFmIEditor.AcUnderLineExecute(Sender: TObject);
begin
  FFontUnderLine := not FFontUnderLine;
  LjnWebBrowserEditor1.SetFontUnderline;
  WBSetFocus;
end;

procedure TFmIEditor.AcBoldExecute(Sender: TObject);
begin
  FFontBold := not FFontBold;
  LjnWebBrowserEditor1.SetFontBold;
  WBSetFocus;
end;

procedure TFmIEditor.AcUpperExecute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetSuperscript;
  WBSetFocus;
end;

procedure TFmIEditor.AcLowerExecute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetSubscript;
  WBSetfocus;
end;

procedure TFmIEditor.LjnWebBrowserEditor1DocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  FDocument:=LjnWebBrowserEditor1.LjnDocument;
  InitFont;
  ClearHTML;
  //AcEdit.Execute;
  FFirst := True;
  EnterEditMode;
end;

procedure TFmIEditor.ClearHTML;
begin
  LjnWebBrowserEditor1.ClearHTML;
  
  //在这里执行一下设置字体的方法，下次输入的字体就是刚才之前选定的字体
  SetFontStyle;
end;

function TFmIEditor.GetHTML: WideString;
begin
  Result := LjnWebBrowserEditor1.LjnOuterHtml;
end;

function TFmIEditor.GetHTMLUtf8: string;
var
  WS: WideString;
begin
{-------------------------------------------------------------------------
  将输出变为 Utf-8 格式。
-------------------------------------------------------------------------}
  WS := LjnWebBrowserEditor1.LjnOuterHtml;
  Result := UTF8Encode(WS);
end;

function TFmIEditor.GetPageCodeStr: string;
var
  CodePage:Word;
begin
  //将当前机器的 CodePage 数字变为字符串
  CodePage := GetACP;
  SetLength(Result,SizeOf(Word));
  Move(CodePage,Result[1],SizeOf(Word));
end;

procedure TFmIEditor.FormCreate(Sender: TObject);
begin
  SetLang;
  
  PNGButton1.Anchors:=PNGButton1.Anchors - [akLeft];
  PNGButton1.Anchors:=PNGButton1.Anchors - [akTop];
  PNGButton1.Anchors:=PNGButton1.Anchors + [akRight];
  PNGButton1.Anchors:=PNGButton1.Anchors + [akBottom];

  ShowToolPanel(True);

  FEnterSend := True;
  N1.Checked := FEnterSend;
  CtrEntertoSend1.Checked := not FEnterSend;

{  AcOpen.Execute;
//  AcOpen.Enabled := False;
  Sleep(10);
  AcEdit.Execute;
//  AcEdit.Enabled := False;
 }
end;

function TFmIEditor.GetHtml2: string;
begin
  Result := LjnWebBrowserEditor1.LjnOuterHtml2;
end;

procedure TFmIEditor.PNGButton1Click(Sender: TObject);
begin
  if Assigned(FOnSendClick) then FOnSendClick(Sender);
end;

procedure TFmIEditor.SetLang;
var
  AName: string;
begin
//  Self.Font.Charset := MyLang.CharSet;

  AName := 'UIEditor';
  AcOpen.Hint := MyLang.LjnGetCaption(AName, 'U101');
  AcEdit.Hint := MyLang.LjnGetCaption(AName, 'U102');
  AcBrowse.Hint := MyLang.LjnGetCaption(AName, 'U103');
  AcSave.Hint := MyLang.LjnGetCaption(AName, 'U104');
  AcCopy.Hint := MyLang.LjnGetCaption(AName, 'U105');
  AcPaste.Hint := MyLang.LjnGetCaption(AName, 'U106');
  AcCut.Hint := MyLang.LjnGetCaption(AName, 'U107');
  AcDelete.Hint := MyLang.LjnGetCaption(AName, 'U108');
  AcLeft.Hint := MyLang.LjnGetCaption(AName, 'U109');
  AcCenter.Hint := MyLang.LjnGetCaption(AName, 'U1010');
  AcRight.Hint := MyLang.LjnGetCaption(AName, 'U111');
  AcHuanHang.Hint := MyLang.LjnGetCaption(AName, 'U112');
  AcLink.Hint := MyLang.LjnGetCaption(AName, 'U113');
  AcLinkDrop.Hint := MyLang.LjnGetCaption(AName, 'U114');
  AcInsertPic.Hint := MyLang.LjnGetCaption(AName, 'U115');
  AcTable.Hint := MyLang.LjnGetCaption(AName, 'U116');
  AcLine.Hint := MyLang.LjnGetCaption(AName, 'U117');
  AcFontColor.Hint := MyLang.LjnGetCaption(AName, 'U118');
  AcFontBgColor.Hint := MyLang.LjnGetCaption(AName, 'U119');
  AcItalic.Hint := MyLang.LjnGetCaption(AName, 'U120');
  AcUnderLine.Hint := MyLang.LjnGetCaption(AName, 'U121');
  AcBold.Hint := MyLang.LjnGetCaption(AName, 'U122');
  AcUpper.Hint := MyLang.LjnGetCaption(AName, 'U123');
  AcLower.Hint := MyLang.LjnGetCaption(AName, 'U124');

  AcCatchScreen.Caption := MyLang.LjnGetHint('3');
  AcCatchScreen.Hint := MyLang.LjnGetHint('3');
  AcCatchScreen2.Caption := MyLang.LjnGetHint('4');

  N1.Caption := MyLang.LjnGetCaption(AName, 'M1');
  CtrEntertoSend1.Caption := MyLang.LjnGetCaption(AName, 'M2');

  //BtnFace.Hint := MyLang.LjnGetCaption(AName, 'U125');
  //BtnEnterSend.Hint := MyLang.LjnGetCaption(AName, 'M1'); //ToDo: 要恢复它
end;

procedure TFmIEditor.InitFont;
begin
  FFontColor := clBlack;
  FFontBgColor := clWhite;
  FFontName := 'Arial';
  FFontSize := 2;
  FFontItalic := False;
  FFontBold := False;
  FFontUnderLine := False;
end;

procedure TFmIEditor.SetFontStyle;
begin
  LjnWebBrowserEditor1.SetFontColor(FFontColor);
  LjnWebBrowserEditor1.SetFontBgColor(FFontBgColor);
  LjnWebBrowserEditor1.SetFontSize(FFontSize);

  if FFontItalic then LjnWebBrowserEditor1.SetFontItalic;
  if FFontBold then LjnWebBrowserEditor1.SetFontBold;
  if FFontUnderLine then LjnWebBrowserEditor1.SetFontUnderline;
  SetFontName(FFontName);
  SetFontSize(FFontSize);
end;

procedure TFmIEditor.LjnWebBrowserEditor1CommandStateChange(
  Sender: TObject; Command: Integer; Enable: WordBool);
begin
  if FFirst then
  begin
    SetFontStyle;
    FFirst := False;
  end;
end;

procedure TFmIEditor.DoEnterUp(Sender: Tobject; Shift: TShiftState;
  Key: word);
begin
  //假设用户设置为按回车发送，则在这里写清除输入框里的文字的代码
  if FEnterSend then
    if Assigned(FOnEnterUp) then FOnEnterUp(self);
end;

procedure TFmIEditor.DoEnterDown(Sender: Tobject; Shift: TShiftState;
  Key: word);
begin
  //假设用户设置为按回车发送，则在这里写发送输入框文字的代码
  //if Assigned(FOnEnterUp) then FOnEnterUp(self);
end;

procedure TFmIEditor.N1Click(Sender: TObject);
begin
  FEnterSend := True;
  N1.Checked := True;
  CtrEntertoSend1.Checked := False;
  LjnWebBrowserEditor1.SetFocus;
  LjnWebBrowserEditor1.SetFocusMe;
end;

procedure TFmIEditor.CtrEntertoSend1Click(Sender: TObject);
begin
  FEnterSend := False;
  N1.Checked := False;
  CtrEntertoSend1.Checked := True;
  LjnWebBrowserEditor1.SetFocus;
  LjnWebBrowserEditor1.SetFocusMe;
end;

procedure TFmIEditor.DoInsertFace(AFileName: string);
begin
  LjnWebBrowserEditor1.DoInsertLocalImage(AFileName);
end;

procedure TFmIEditor.FaceSelectMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APoint, BPoint: TPoint;
begin
  APoint.X := X;
  APoint.Y := Y;

  BPoint := TControl(Sender).ClientToScreen(APoint);

  if not Assigned(FmSelectFace) then FmSelectFace := TFmSelectFace.Create(Application);
  FmSelectFace.Left := BPoint.X - X;
  FmSelectFace.Top := BPoint.Y - Y - FmSelectFace.Height;
  FmSelectFace.NLOnSelectPic := DoInsertFace;
  FmSelectFace.Show;
end;

procedure TFmIEditor.FaceSelectMouseLeave(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetFocus;
  LjnWebBrowserEditor1.SetFocusMe;
end;

procedure TFmIEditor.EnterMenuMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APoint, BPoint: TPoint;
begin
  APoint.X := X;
  APoint.Y := Y;

  BPoint := TControl(Sender).ClientToScreen(APoint);
  PopupMenu1.Popup(BPoint.X, BPoint.Y);
end;

procedure TFmIEditor.ShowToolPanel(AStatus: Boolean);
begin
  if AStatus then
    Panel1.Height := 30
  else Panel1.Height := 60;
end;

procedure TFmIEditor.PNGButton28Click(Sender: TObject);
begin
  if Panel1.Height = 30 then Panel1.Height := 60 else Panel1.Height := 30;
end;


procedure TFmIEditor.OpenBrowser;
begin
  LjnWebBrowserEditor1.OpenBrowser;
  LjnWebBrowserEditor1.LjnOnKeyDown := DoOnKeyDown;
  LjnWebBrowserEditor1.LjnOnAlt_S_DownEvent := DoAltS;
  LjnWebBrowserEditor1.LjnOnCtr_Enter_DownEvent := DoCtrEnter;
  LjnWebBrowserEditor1.LJnOnEnter_UpEvent := DoEnterUp;
  LjnWebBrowserEditor1.LjnOnEnter_DownEvent := DoEnterDown;
end;

procedure TFmIEditor.AcCatchScreenExecute(Sender: TObject);
var
  APath: string;
begin
  APath := NLUserInfo.NLRootPath + NL_CatchSCreen_Dir;
  if not DirectoryExists(APath) then
    if not CreateDir(APath) then APath := NLUserInfo.NLRootPath;

  FmFullScreen.NLPicFilePath := APath;
  FmFullScreen.NLAfterCatchScreen := DoAfterCatchScreen;
  FmFullScreen.Go;
end;

procedure TFmIEditor.DoAfterCatchScreen(AFileName: string);
begin
  FmFullScreen.Close;
  if FileExists(AFileName) then
  begin
    //把图插入编辑器
    LjnWebBrowserEditor1.DoInsertLocalImage(AFileName);
  end;
end;

procedure TFmIEditor.AcCatchScreen2Execute(Sender: TObject);
var
  APath: string;
begin
  //让会议窗口缩到最小，不阻挡屏幕
  if Assigned(FOnMinMeetingForm) then FOnMinMeetingForm(Self);
  Sleep(50);
  APath := NLUserInfo.NLRootPath + NL_CatchSCreen_Dir;
  if not DirectoryExists(APath) then
    if not CreateDir(APath) then APath := NLUserInfo.NLRootPath;

  FmFullScreen.NLPicFilePath := APath;
  FmFullScreen.NLAfterCatchScreen := DoAfterCatchScreen;

  FmFullScreen.Go;

  //FmFullScreen.Show; //要让这个 Form 获得焦点。否则可能就是会议 form 获得焦点了。

  if Assigned(FOnRestoreMeetingForm) then FOnRestoreMeetingForm(self); //还原会议窗口
  FmFullScreen.Show; //要让这个 Form 获得焦点。否则可能就是会议 form 获得焦点了。
end;

procedure TFmIEditor.PNGButton29MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  APoint, BPoint: TPoint;
begin
  if Button = mbRight then
  begin
    APoint.X := X;
    APoint.Y := Y;

    BPoint := TControl(Sender).ClientToScreen(APoint);

    PopupMenu2.Popup(BPoint.X - X, BPoint.Y - Y);
  end;
  //FmSelectFace.Left := BPoint.X - X;
  //FmSelectFace.Top := BPoint.Y - Y - FmSelectFace.Height;
end;

procedure TFmIEditor.RzBackground1DblClick(Sender: TObject);
var
  i: Integer;
begin
  //RzPanel1.Visible := not RzPanel1.Visible;
  //RzPanel2.Visible := RzPanel1.Visible;

  for i := 0 to self.ComponentCount -1 do
  begin
    if (self.Components[i] is TRzToolButton) then
    begin
      TRzToolButton(self.Components[i]).Visible := not TRzToolButton(self.Components[i]).Visible;
    end;
  end;
end;

end.
