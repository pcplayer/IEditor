unit UFmLjnEditor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, ImgList, StdActns, ExtActns,
  ActnMan, ActnCtrls, ToolWin, ActnMenus,
  ExtCtrls, OleCtrls, SHDocVw, WBrowserEditor, StdCtrls,
  {RzCmboBx, RibbonActnCtrls, RzPanel, Mask, RzEdit, RzSpnEdt, RzButton,}
  PlatformDefaultStyleActnCtrls, JvExStdCtrls, JvCombobox,  JvSpin, JvColorCombo,
  Vcl.Mask, JvExMask, JvExControls, JvButton, JvTransparentButton, {UIrBPLFactoryIntf,}
  System.Actions, System.ImageList;

type
  TFmLjnEditor = class(TForm{, IIeEditor}) //这个接口，将来用于包编译时用。暂时现在先屏蔽掉。2013-3-4
    ActionManager1: TActionManager;
    FileOpen1: TFileOpen;
    FileSaveAs1: TFileSaveAs;
    FilePrintSetup1: TFilePrintSetup;
    FilePageSetup1: TFilePageSetup;
    FileExit1: TFileExit;
    EditCut1: TEditCut;
    EditCopy1: TEditCopy;
    EditPaste1: TEditPaste;
    EditSelectAll1: TEditSelectAll;
    EditUndo1: TEditUndo;
    EditDelete1: TEditDelete;
    FileSaveActn: TAction;
    FileNewActn: TAction;
    RichEditBold1: TRichEditBold;
    RichEditItalic1: TRichEditItalic;
    FontCalibriActn: TAction;
    FontCambriaActn: TAction;
    FontCourierNewActn: TAction;
    FontArialRoundedMTBoldActn: TAction;
    FontArialActn: TAction;
    FontArialNarrowActn: TAction;
    FontTahomaActn: TAction;
    FontSegoeUIActn: TAction;
    FontSegoeScriptActn: TAction;
    RichEditUnderline1: TRichEditUnderline;
    RichEditStrikeOut1: TRichEditStrikeOut;
    RichEditBullets1: TRichEditBullets;
    RichEditAlignLeft1: TRichEditAlignLeft;
    RichEditAlignRight1: TRichEditAlignRight;
    RichEditAlignCenter1: TRichEditAlignCenter;
    SearchFind1: TSearchFind;
    SearchReplace1: TSearchReplace;
    LunaStyleActn: TAction;
    ObsidianStyleActn: TAction;
    SilverStyleActn: TAction;
    FileCloseActn: TAction;
    FileSaveAsText: TAction;
    FileSaveAsRTF: TAction;
    FileQuickPrint: TAction;
    FilePrintPreview: TAction;
    FontGrowSizeActn: TAction;
    FontShrinkSizeActn: TAction;
    FontSubscriptActn: TAction;
    FontSuperScriptActn: TAction;
    ChangeCaseSentenceActn: TAction;
    ChangeCaseLowerActn: TAction;
    ChangeCaseUpperActn: TAction;
    ChangeCaseCapitalizeActn: TAction;
    ChangeCaseToggleActn: TAction;
    FontHighlightActn: TAction;
    FontColorActn: TAction;
    ChangeCaseActn: TAction;
    EditPasteSpecial: TAction;
    EditPasteHyperlink: TAction;
    FileRun1: TFileRun;
    FontEdit1: TFontEdit;
    PrintDlg1: TPrintDlg;
    RadioAction1: TAction;
    RadioAction2: TAction;
    RadioAction3: TAction;
    CheckboxAction1: TAction;
    CheckboxAction2: TAction;
    CheckboxAction3: TAction;
    NumberingActn: TAction;
    ilGFX16: TImageList;
    ilGFX32: TImageList;
    ilGFX16_d: TImageList;
    ilGFX32_d: TImageList;
    alBulletNumberGallery: TActionList;
    NumberNoneActn: TAction;
    NumberArabicDotActn: TAction;
    NumberArabicParenActn: TAction;
    NumberUpperRomanActn: TAction;
    NumberUpperActn: TAction;
    NumberLowerParenActn: TAction;
    NumberLowerDotActn: TAction;
    NumberLowerRomanActn: TAction;
    ilBulletNumberGallery: TImageList;
    ColorDialog1: TColorDialog;
    AcFontBackColor: TAction;
    AcPicAdd: TAction;
    AcScreenCatch: TAction;
    AcHideToolPanel: TAction;
    Panel1: TPanel;
    RzPanel1: TPanel;
    RzToolButton1: TJvTransparentButton;
    RzToolButton2: TJvTransparentButton;
    RzToolButton3: TJvTransparentButton;
    RzToolButton4: TJvTransparentButton;
    RzToolButton5: TJvTransparentButton;
    FontComboBox1: TJvFontComboBox;
    RzSpinEdit1: TJvSpinEdit;
    RzPanel2: TPanel;
    RzToolButton6: TJvTransparentButton;
    RzToolButton7: TJvTransparentButton;
    RzToolButton8: TJvTransparentButton;
    RzToolButton9: TJvTransparentButton;
    RzToolButton10: TJvTransparentButton;
    RzToolButton11: TJvTransparentButton;
    AcAddLink: TAction;
    AcDeleteLink: TAction;
    RzToolButton12: TJvTransparentButton;
    AcInsertTable: TAction;
    RzToolButton13: TJvTransparentButton;
    AcInsertLine: TAction;
    RzToolButton14: TJvTransparentButton;
    RzToolButton15: TJvTransparentButton;
    RzToolButton16: TJvTransparentButton;
    RzToolButton17: TJvTransparentButton;
    AcInsertParagraph: TAction;
    RzToolButton18: TJvTransparentButton;
    RzPanel3: TPanel;
    RzPanel4: TPanel;

    RzToolButton19: TJvTransparentButton;
    RzToolButton20: TJvTransparentButton;
    RzToolButton21: TJvTransparentButton;
    RzToolButton22: TJvTransparentButton;
    procedure EditCut1Execute(Sender: TObject);
    procedure EditCopy1Execute(Sender: TObject);
    procedure EditPaste1Execute(Sender: TObject);
    procedure EditDelete1Execute(Sender: TObject);
    procedure FontComboBox1Change(Sender: TObject);
    procedure RzSpinEdit1Change(Sender: TObject);
    procedure FontColorActnExecute(Sender: TObject);
    procedure AcFontBackColorExecute(Sender: TObject);
    procedure AcPicAddExecute(Sender: TObject);
    procedure AcScreenCatchExecute(Sender: TObject);
    procedure AcHideToolPanelExecute(Sender: TObject);
    procedure RichEditBold1Execute(Sender: TObject);
    procedure RichEditItalic1Execute(Sender: TObject);
    procedure RichEditUnderline1Execute(Sender: TObject);
    procedure FontSuperScriptActnExecute(Sender: TObject);
    procedure FontSubscriptActnExecute(Sender: TObject);
    procedure AcAddLinkExecute(Sender: TObject);
    procedure AcDeleteLinkExecute(Sender: TObject);
    procedure AcInsertTableExecute(Sender: TObject);
    procedure AcInsertLineExecute(Sender: TObject);
    procedure RichEditAlignLeft1Execute(Sender: TObject);
    procedure RichEditAlignCenter1Execute(Sender: TObject);
    procedure RichEditAlignRight1Execute(Sender: TObject);
    procedure AcInsertParagraphExecute(Sender: TObject);
    procedure RichEditStrikeOut1Execute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    FFontName: string;
    FFontSize: Integer;
    FFontColor: TColor;
    FToolPanelStatus: Boolean;
    LjnWebBrowserEditor1: TLjnWebBrowserEditor;

    procedure WBSetFocus;
    procedure SetFontName(const FontName: string);
    procedure SetFontSize(const ASize: Integer);
    procedure DoAfterCatchScreen(AFileName: string);
    procedure ShowToolPanel(AStatus: Boolean);
    function GetHTML: WideString;
    procedure SetHTML(const Value: WideString);
    function GetReadyState: WideString;
    function GetTempPath: string;
    procedure SetTempPath(const Value: string);
    function GetContentTransfer: string;
    procedure SetContentTransfer(const Value: string);
  public
    { Public declarations }
    procedure OpenEditor;
    procedure SetBorderNone;
    procedure LoadMIMEFromStream(AStream:TStream; AppendData: Boolean; CodePage:Word);  //AppendData=True 则保留原来的内容，新内容追加在后面
    procedure ClearHTML;

    property HTMLEncoded: WideString read GetHTML write SetHTML;
    property LjnReadyState: WideString read GetReadyState;
    property LjnTempPath: string read GetTempPath write SetTempPath;
    property LjnContentTransfer: string read GetContentTransfer write SetContentTransfer;
  end;

var
  FmLjnEditor: TFmLjnEditor;

implementation

uses UFmFullScreen, UTableProperty;

{$R *.dfm}

{ TFmLjnEditor }

procedure TFmLjnEditor.AcAddLinkExecute(Sender: TObject);
begin
  //插入链接
  LjnWebBrowserEditor1.DoCreateLink;
end;

procedure TFmLjnEditor.AcDeleteLinkExecute(Sender: TObject);
begin
  //去掉链接
  LjnWebBrowserEditor1.DoUnlink;
end;

procedure TFmLjnEditor.AcFontBackColorExecute(Sender: TObject);
begin
  if ColorDialog1.Execute then
    LjnWebBrowserEditor1.SetFontBgColor(ColorDialog1.Color);
end;

procedure TFmLjnEditor.AcHideToolPanelExecute(Sender: TObject);
begin
{-------------------------------------------------------------------------------
  显示/隐藏多余的按钮栏
-------------------------------------------------------------------------------}
  ShowToolPanel(not FToolPanelStatus);
end;

procedure TFmLjnEditor.AcInsertLineExecute(Sender: TObject);
begin
  //插入横线
  LjnWebBrowserEditor1.DoInsertHorizontalRule;
end;

procedure TFmLjnEditor.AcInsertParagraphExecute(Sender: TObject);
begin
  //换段落
  LjnWebBrowserEditor1.SetInsertParagraph;
end;

procedure TFmLjnEditor.AcInsertTableExecute(Sender: TObject);
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

procedure TFmLjnEditor.AcPicAddExecute(Sender: TObject);
begin
  LjnWebBrowserEditor1.DoInsertImage;
end;

procedure TFmLjnEditor.AcScreenCatchExecute(Sender: TObject);
begin
  if not Assigned(FmFullScreen) then
    FmFullScreen := TFmFullScreen.Create(Application);

  FmFullScreen.NLAfterCatchScreen := DoAfterCatchScreen;
  FmFullScreen.Go;
end;

procedure TFmLjnEditor.ClearHTML;
begin
  LjnWebBrowserEditor1.ClearHTML;
end;

procedure TFmLjnEditor.DoAfterCatchScreen(AFileName: string);
begin
  FmFullScreen.Close;
  if FileExists(AFileName) then
  begin
    //把图插入编辑器
    LjnWebBrowserEditor1.DoInsertLocalImage(AFileName);
  end;
end;

procedure TFmLjnEditor.EditCopy1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.DoCopy;
end;

procedure TFmLjnEditor.EditCut1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.DoCut;
  WBSetFocus;
end;

procedure TFmLjnEditor.EditDelete1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.DoDelete;
end;

procedure TFmLjnEditor.EditPaste1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.DoPaste;
end;

procedure TFmLjnEditor.OpenEditor;
begin
  LjnWebBrowserEditor1.OpenBrowser;
  LjnWebBrowserEditor1.LjnEditModeOn := True;
end;

procedure TFmLjnEditor.RichEditAlignCenter1Execute(Sender: TObject);
begin
  //居中
  LjnWebBrowserEditor1.SetJustifyCenter;
end;

procedure TFmLjnEditor.RichEditAlignLeft1Execute(Sender: TObject);
begin
  //左对齐
  LjnWebBrowserEditor1.SetJustifyLeft;
end;

procedure TFmLjnEditor.RichEditAlignRight1Execute(Sender: TObject);
begin
  //右对齐
  LjnWebBrowserEditor1.SetJustifyRight;
end;

procedure TFmLjnEditor.RichEditBold1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetFontBold;
end;

procedure TFmLjnEditor.RichEditItalic1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetFontItalic;
end;

procedure TFmLjnEditor.RichEditStrikeOut1Execute(Sender: TObject);
begin
  //删除线
  LjnWebBrowserEditor1.SetFontStrikeThrough;
end;

procedure TFmLjnEditor.RichEditUnderline1Execute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetFontUnderline;
end;

procedure TFmLjnEditor.RzSpinEdit1Change(Sender: TObject);
var
  i: Integer;
begin
  i := Trunc(RzSpinEdit1.Value);
  Self.SetFontSize(i);
end;

procedure TFmLjnEditor.SetBorderNone;
begin
  LjnWebBrowserEditor1.SetBorderNone;
end;

procedure TFmLjnEditor.SetContentTransfer(const Value: string);
begin
  LjnWebBrowserEditor1.LjnContentTransfer := Value;
end;

procedure TFmLjnEditor.SetFontName(const FontName: string);
begin
  FFontName := FontName;
  LjnWebBrowserEditor1.SetFontName(FontName);
end;

procedure TFmLjnEditor.SetFontSize(const ASize: Integer);
begin
  FFontSize := ASize;
  LjnWebBrowserEditor1.SetFontSize(ASize);
end;

procedure TFmLjnEditor.SetHTML(const Value: WideString);
begin
  //todo:
end;

procedure TFmLjnEditor.SetTempPath(const Value: string);
begin
  LjnWebBrowserEditor1.LjnTempPath := Value;
end;

procedure TFmLjnEditor.ShowToolPanel(AStatus: Boolean);
begin
  if AStatus then
  begin
    Panel1.Height := 30;
    RzPanel2.Visible := False;
    //RzPanel3.Visible := RzPanel2.Visible;
  end
  else
  begin
    Panel1.Height := 60;
    RzPanel2.Visible := True;
    //RzPanel3.Visible := RzPanel2.Visible;
  end;

  FToolPanelStatus := AStatus;
end;

procedure TFmLjnEditor.FontColorActnExecute(Sender: TObject);
begin
  if ColorDialog1.Execute then
  begin
    FFontColor := ColorDialog1.Color;
    LjnWebBrowserEditor1.SetFontColor(ColorDialog1.Color);
    WBSetFocus;
  end;
end;

procedure TFmLjnEditor.FontComboBox1Change(Sender: TObject);
begin
  SetFontName(FontComboBox1.FontName);
end;

procedure TFmLjnEditor.FontSubscriptActnExecute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetSubscript;
end;

procedure TFmLjnEditor.FontSuperScriptActnExecute(Sender: TObject);
begin
  LjnWebBrowserEditor1.SetSuperscript;
end;

procedure TFmLjnEditor.FormCreate(Sender: TObject);
var
  i: Integer;
begin
  Self.ShowToolPanel(True);
  if not Assigned(LjnWebBrowserEditor1) then
  begin
    LjnWebBrowserEditor1 := TLjnWebBrowserEditor.Create(Self);
    //LjnWebBrowserEditor1.Parent := RzPanel4;
    LjnWebBrowserEditor1.Align := alClient;
    LjnWebBrowserEditor1.SetParentComponent(RzPanel4);
    LjnWebBrowserEditor1.Show;
  end;

  for i := 0 to Self.ComponentCount -1 do
  begin
    if (Self.Components[i] is TJvTransparentButton) then
    begin
      TJvTransparentButton(Self.Components[i]).Images.ActiveIndex := TAction(TJvTransparentButton(Self.Components[i]).Action).ImageIndex;
      TJvTransparentButton(Self.Components[i]).Caption := '';
      TJvTransparentButton(Self.Components[i]).Images.ActiveImage := ilGFX16;
    end;

  end;

end;

function TFmLjnEditor.GetContentTransfer: string;
begin
  Result := LjnWebBrowserEditor1.LjnContentTransfer;
end;

function TFmLjnEditor.GetHTML: WideString;
begin
  Result := LjnWebBrowserEditor1.LjnOuterHtml3;
end;

function TFmLjnEditor.GetReadyState: WideString;
begin
  Result := LjnWebBrowserEditor1.LjnReadyState; //浏览器是否准备好
end;

function TFmLjnEditor.GetTempPath: string;
begin
  Result := LjnWebBrowserEditor1.LjnTempPath;
end;

procedure TFmLjnEditor.LoadMIMEFromStream(AStream: TStream; AppendData: Boolean;
  CodePage: Word);
begin
  //LjnWebBrowserEditor1.LoadMIMEFromStream(AStream, AppendData, CodePage);
end;

procedure TFmLjnEditor.WBSetFocus;
begin
  Self.Show;
  LjnWebBrowserEditor1.SetFocus;
  LjnWebBrowserEditor1.SetFocusMe;
end;


initialization
  //UIrBPLFactoryIntf.RegClassToIrFactory(TFmLjnEditor);

end.
