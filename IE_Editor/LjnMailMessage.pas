unit LjnMailMessage;
{-----------------------------------------------------------------------
  �����ʼ����ࡣ������ʼ���Ϣ���/���뵽 TStream �� File �
  ʹ�ñ���: �����ʼ��������ʼ����ѽ��յ����ʼ���ԭ���ı��뱣�浽���ݿ��
  �����ݿ���ȡ���ʼ���������������������ʾ����
  ��ʾ��ʱ������� HTML ��ʽ������ֻ������ı���Ҳ������� HTML.
  ��ʾʱ�������ͼƬ����ͼƬ��������ʱĿ¼������ͼƬ�����޸�Ϊ��ʱĿ¼����ļ���
  ���ݿ������ֶα����������ʼ���CharSet/ContentTransferEncoding/ContentType/MsgId/UID ��

  ������Stream��File��ȡ�ʼ����ݣ���������¼��������ͳ�����
----------------------------------------------------------------------------}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,IdMessage,
   ComCtrls, IdSMTP, IdComponent, IdTCPConnection,Contnrs,
   IdTCPClient, {IdMessageClient,} IdPOP3, IdBaseComponent, IdAttachment, IdText,
   {IdAntiFreezeBase,} IdMessageCoder, IdMessageCoderMIME,
   LjnMiMeDecode, UMailCommon;

type
  TLjnPOP3=class(TObject)
  private
    FPOP3:TIdPOP3;
  public
  end;

  TAttachParam=class(TObject)
  public
    FFileName : string;
    FFileSize : string;
    ContentType : string;
    ContentID : string;
  end;

  TMailText=class(TObject)
  public
    FText:string;
    ContentType:string;
  end;

  TAttachParamList=class(TObject)
  private
    FList:TObjectList;
    function GetCount: Integer;
    function GetItem(Index: Integer): TAttachParam;
    procedure SetItem(Index: Integer; const Value: TAttachParam);
  public
    constructor Create;
    destructor Destroy;override;

    procedure Add(Attach:TAttachParam);
    procedure Clear;

    property Count:Integer read GetCount;
    property Items[Index: Integer]: TAttachParam read GetItem write SetItem; default;
  end;

  TMailTextList=class(TObject)
  private
    FList:TObjectList;
    function GetCount: Integer;
    function GetItem(Index: Integer): TMailText;
    procedure SetItem(Index: Integer; const Value: TMailText);
    function GetHtmlText: Widestring;
    function GetIsHTML: Boolean;
    function GetPlainText: WideString;
  public
    constructor Create;
    destructor Destroy;override;

    procedure Add(MailText:TMailText);
    procedure Clear;

    property Count:Integer read GetCount;
    property LjnPlainText:WideString read GetPlainText;
    property LjnHtmlText:Widestring read GetHtmlText;
    property LjnIsHTML:Boolean read GetIsHTML;
    property Items[Index: Integer]: TMailText read GetItem write SetItem; default;
  end;

  //��һ�㣬�����ֻ���� Message �������շ�
  //�����ר�Źܽ��롣����ǰ�����������ʼ������������ݿ���ļ�����������޹ء�
  TAfterMailDecode = procedure(MailContent:TMailTextList; AttachList:TAttachParamList) of object;
  
  TLjnMailMessage=class(TObject)
  private
    FMsg:TIdMessage;
    FLjnMIMEDecode:TLjnMIMEDecode;

    FIsMultiPart:Boolean;
    FTmpPath:string;

    FAttachParamList:TAttachParamList;
    FMailTextList:TMailTextList;

    //�¼�
    FAfterMailDecode:TAfterMailDecode;

    function GetHtmlText: WideString;
    function GetIsHTML: Boolean;
    function GetPlainText: Widestring;
    function GetCC: string;
    function GetDate: TDateTime;
    function GetFrom: string;
    function GetReceipt: string;
    function GetRecipients: string;
    function GetSubject: string;

    function GetFileNameByCID(CID:string):string;

    procedure GetMailContent;
    function GetTmpPath: string;
    procedure SetTmpPath(const Value: string);
    
    procedure ReplaceContentIDInMailText(var S:string);
    procedure DoAfterMailDecode;
  public
    constructor Create;
    destructor Destroy;override;
    
    procedure LoadFromFile(AFileName:string);
    procedure LoadFromStream(AStream:TStream);

    property LjnPlainText:WideString read GetPlainText;
    property LjnHtmlText:Widestring read GetHtmlText;
    property LjnIsHTML:Boolean read GetIsHTML;
    property LjnTmpPath:string read GetTmpPath write SetTmpPath;

    property LjnFrom:string read GetFrom;
    property LjnRecipients:string read GetRecipients;
    property LjnSubject:string read GetSubject;
    property LjnDate:TDateTime read GetDate;
    property LjnCC:string read GetCC;
    property LjnReceipt:string read GetReceipt;


    //�¼�
    property AfterMailDecode:TAfterMailDecode read FAfterMailDecode write FAfterMailDecode;
  end;

implementation

{ TLjnMailMessage }

constructor TLjnMailMessage.Create;
begin
  FMsg:=TIdMessage.Create(nil);

  FMsg.NoEncode:=False;
  FMsg.NoDecode:=False;

  FLjnMIMEDecode:=TLjnMIMEDecode.Create;

  FAttachParamList:=TAttachParamList.Create;
  FMailTextList:=TMailTextList.Create;

  {
  SetLength(FTmpPath, 256);
  GetTempPathA(255, PChar(FTmpPath));
  Trim(FTmpPath);
  }
end;

destructor TLjnMailMessage.Destroy;
begin
  FMsg.Free;
  FLjnMimeDecode.Free;
  FAttachParamList.Free;
  FMailTextList.Free;
  inherited;
end;

procedure TLjnMailMessage.ReplaceContentIDInMailText(var S: string);
type
  TFindStatus=(mtsIdle,mtsLeft,mtsRight,mtsI,mtsM,mtsG,mtsS,mtsR,mtsC,mtsEqu,
    mtsQuoteOne,mtsCIDc,mtsCIDi,mtsCIDd,mtsCID);
var
  i,CidBegin,CidEnd:Integer;
  St:TFindStatus;
  S1,CidStr,AFileName:string;
begin
{-------------------------------------------------------------------------
  ʹ��״̬�������ʼ��ı����ҳ������˸�����ͼ�ţ���ContentID�ĵط���
  ���͵��ַ������ӣ�(Ƭ��)
  <DIV><IMG alt="" hspace=0 src="cid:__0@Foxmail.net" align=baseline
  border=0></DIV>
  <DIV>&nbsp;</DIV>
------------------------------------------------------------------------------}

  St:=mtsIdle;
  CidStr:='';
  if Length(S)<1 then Exit;

  i:=1;

  while True do
  begin
    case St of
      mtsIdle:
      begin
        if S[i]='<' then St:=mtsLeft;
      end;

      mtsLeft:
      begin
        if ((S[i]='I') or (S[i]='i')) then St:=mtsI else St:=mtsIdle;
      end;

      mtsRight:
      begin

      end;

      mtsI:
      begin
        if ((S[i]='M') or (S[i]='m')) then St:=mtsM else St:=mtsIdle;
      end;

      mtsM:
      begin
        if ((S[i]='G') or (S[i]='g')) then St:=mtsG else St:=mtsIdle;
      end;

      mtsG:
      begin
        if ((S[i]='S') or (S[i]='s')) then St:=mtsS;
      end;

      mtsS:
      begin
        if ((S[i]='R') or (S[i]='r')) then St:=mtsR else St:=mtsG;
      end;

      mtsR:
      begin
        if ((S[i]='C') or (S[i]='c')) then St:=mtsC else St:=mtsIdle;
      end;

      mtsC:
      begin
        if (S[i]='=') then St:=mtsEqu else St:=mtsIdle;
      end;

      mtsEqu:
      begin
        if ((S[i]='"') or (S[i]='''')) then St:=mtsQuoteOne else St:=mtsIdle;
      end;

      mtsQuoteOne:
      begin
        if ((S[i]='C') or (S[i]='c')) then
        begin
          St:=mtsCIDc;
          CidBegin:=i;
        end
        else St:=mtsIdle;
      end;

      mtsCIDc:
      begin
        if ((S[i]='I') or (S[i]='i')) then St:=mtsCIDI
        else
        begin
          St:=mtsIdle;
          CidBegin:=0;
        end;
      end;

      mtsCIDi:
      begin
        if ((S[i]='D') or (S[i]='d')) then St:=mtsCIDd
        else
        begin
          St:=mtsIdle;
          CidBegin:=0;
        end;
      end;

      mtsCIDd:
      begin
        if (S[i]=':') then St:=mtsCID
        else
        begin
          St:=mtsIdle;
          CidBegin:=0;
        end;
      end;

      mtsCID:
      begin
        if S[i]<>'"' then
        begin
          CidStr:=CidStr+S[i];
        end
        else
        begin
          CidEnd:=i-1;

          //���뺯������CID �滻Ϊ��Ӧ���ļ����������Ϳ���һ�ΰ������ı������꣬�����ظ�����
          AFileName:=GetFileNameByCID(CidStr);
          if AFileName<>'' then
          begin
            //ȥ���ı��м��CID
            Delete(S,CidBegin,CidEnd-CidBegin+1);

            //�� CidBegin ��λ�ð��ļ�������
            Insert(AFileName,S,CidBegin);
          end;
          
          St:=mtsIdle;
          CidBegin:=0;
          CidEnd:=0;
          CidStr:='';
        end;
      end;
    end; //case

    Inc(i);
    if i>Length(S) then Break;
  end;
end;

function TLjnMailMessage.GetCC: string;
var
  WS:WideString;
begin
  FLjnMimeDecode.DecodeMIME(FMsg.CCList.EMailAddresses,WS);
  Result:=WS;
end;

function TLjnMailMessage.GetDate: TDateTime;
begin
  Result:=FMsg.Date;
end;

function TLjnMailMessage.GetFileNameByCID(CID: string): string;
var
  i:Integer;
  MyAttach:TAttachParam;
begin
  //���� EMail �ı��е� CID��ȡ�ò����ڸ�λ�õ��ļ����ļ��������ǶԲ���ͼƬ�� HTML ��ʽ���ʼ�������˵�ġ�
  Result:='';
  for i:=0 to FAttachParamList.Count-1 do
  begin
    MyAttach:=FAttachParamList.Items[i];
    //if Pos(CID,MyAttach.ContentID)>0 then
    if Pos(UpperCase(CID),UpperCase(MyAttach.FFileName))>0 then
    begin
      Result:=MyAttach.FFileName;
      Break;
    end;
  end;

  if Result<>'' then
  begin
    Result:=FtmpPath+Result;
  end;
end;

function TLjnMailMessage.GetFrom: string;
var
  WS:WideString;
begin
  FLjnMimeDecode.DecodeMIME(FMsg.From.Text,WS);
  Result:=WS;
end;

function TLjnMailMessage.GetHtmlText: Widestring;
var
  i:Integer;
  MailText:TMailText;
  S:WideString;
begin
  {S:='';
  for i:=0 to Pred(FMailTextList.Count) do
  begin
    MailText:=FMailTextList.Items[i];
    if POS('HTML', UpperCase(MailText.ContentType))>0 then
    begin
      S:=S+MailText.FText;
    end;
  end;
  Result:=S;
  }
  Result:=FMailTextList.LjnHtmlText;
end;

function TLjnMailMessage.GetIsHTML: Boolean;
var
  i:Integer;
  MailText:TMailText;
begin
{  Result:=False;
  for i:=0 to Pred(FMailTextList.Count) do
  begin
    MailText:=FMailTextList.Items[i];
    if POS('HTML', UpperCase(MailText.ContentType))>0 then
    begin
      Result:=True;
      Break;
    end;
  end;
  }
  Result:=FMailTextList.LjnIsHTML;
end;

procedure TLjnMailMessage.GetMailContent;
var
  i,j,AFileSize:Integer;
  AFileName:WideString;
  S,ContentS:string;
  ContentWS : WideString;
  AttachParam:TAttachParam;
  AMailText:TMailText;
  GUID: TGUID;
begin
  //ȡ�� Mail �����ݡ�����Ǳ���ģ�����롣�����ͼƬ����ͼƬд����ʱĿ¼��
  FMailTextList.Clear;
  FAttachParamList.Clear;
  if Pos('PLAIN', UpperCase(FMsg.ContentType))>0 then
  begin
    AMailText:=TMailText.Create;
    AMailText.FText:=FLjnMimeDecode.DecodeMailBody(FMsg.ContentTransferEncoding,FMsg.CharSet,FMsg.ContentType,FMsg.Body.Text,FMsg.Encoding);
    AMailText.ContentType:=UpperCase(FMsg.ContentType);
    FMailTextList.Add(AMailText);
  end;

   //����ж������
  for i:=0 to Pred(FMsg.MessageParts.Count) do
  begin
    if (FMsg.MessageParts.Items[i] is TIdAttachment) then
    begin
      if (FMsg.MessageParts.Items[i] is TIdAttachment) then
      begin
        AttachParam:=TAttachParam.Create;
        S:=TIdAttachment(FMsg.MessageParts.Items[i]).Filename;
        AFileName := '';
        FLjnMimeDecode.DecodeMIME(S,AFileName);  //�ļ���Ҳ�п����Ǳ����
        AttachParam.FFileName := AFileName;
        AttachParam.ContentType := TIdAttachment(FMsg.MessageParts.Items[i]).ContentType;
        AttachParam.ContentID := TIdAttachment(FMsg.MessageParts.Items[i]).ContentID;

        if FTmpPath='' then raise Exception.Create('���渽�����ļ���û������!');

        //CreateGUID(GUID);

        //AFileName := GUIDToString(GUID)+ ExtractFileExt(AFileName); //add by pcplayer 2007-8-21
        if FileExists(FTmpPath+AFileName) then
        begin
          if not DeleteFile(FTmpPath+AFileName) then raise Exception.Create('��ͬ���ļ�����ɾ��!');
        end;
        TIdAttachment(FMsg.MessageParts.Items[i]).SaveToFile(FTmpPath+AFileName);

        AFileSize := GetFileSizeLjn(FTmpPath+AFileName); 
        AttachParam.FFileSize := FileSizeToStr(AFileSize);

        FAttachParamList.Add(AttachParam);
      end;
    end
    else
    begin
      if FMsg.MessageParts.Items[i] is TIdText then
      begin
         {-------------------------------------------------------------------------------
            ��� ContentType = TEXT/XML ����������Լ������ݿ⣡
         ------------------------------------------------------------------------------------}
         AMailText:=TMailText.Create;
         AMailText.FText:=TIdText(FMsg.MessageParts.Items[i]).Body.Text;
         AMailText.ContentType:=UpperCase(TIdText(FMsg.MessageParts.Items[i]).ContentType);
         FMailTextList.Add(AMailText);
      end;
    end;
  end;

  if FMsg.MessageParts.Count=0 then
  begin
    AMailText:=TMailText.Create;
    ContentS := FMsg.Body.Text;

    { Indy 10 �£�ֻҪ ContentTransferEncoding �� base64 ���Զ������ˡ�����Ҫ���ô�����롣
    if UpperCase(FMsg.ContentTransferEncoding) = 'BASE64' then    //ToDo: ���� base64�������������ֱ��룬Ҫ�����������email ��������ַ���
    begin
      //Ҫ�����ݽ��н���
      ContentWS := FLjnMIMEDecode.DecodeBase64(ContentS);
      AMailText.FText := ContentWS;
    end
    else AMailText.FText := ContentS;
    }

    AMailText.FText := ContentS; //add for Indy 10

    AMailText.ContentType := FMsg.ContentType;
    if AMailText.ContentType='' then
      AMailText.ContentType:='PLAIN/TEXT';
    FMailTextList.Add(AMailText);
  end;

  //�ҳ� HTML ������У����ҳ������ͼƬ���滻Ϊ�ļ���
  for i:=0 to FMailTextList.Count-1 do
  begin
    AMailText:=FMailTextList.Items[i];
    if Pos('TEXT/HTML', UpperCase(AMailText.ContentType))>0 then
    begin
      S:=AMailText.FText;
      ReplaceContentIDInMailText(S); //��������в����ͼƬ(CID)�����ҵ�cid���滻Ϊ��������ļ���
      AMailText.FText := S;
    end;
  end;

  //������ϣ������¼�
  DoAfterMailDecode;
end;

function TLjnMailMessage.GetPlainText: WideString;
var
  i:Integer;
  MailText:TMailText;
  S:WideString;
begin
  {S:='';
  for i:=0 to Pred(FMailTextList.Count) do
  begin
    MailText:=FMailTextList.Items[i];
    if POS('PLAIN', UpperCase(MailText.ContentType))>0 then
    begin
      S:=S+MailText.FText;
    end;
  end;
  Result:=S;
  }
  Result := FMailTextList.LjnPlainText;
end;

function TLjnMailMessage.GetReceipt: string;
var
  WS:WideString;
begin
  FLjnMimeDecode.DecodeMIME(FMsg.ReceiptRecipient.Text,WS);
  Result:=WS;
end;

function TLjnMailMessage.GetRecipients: string;
var
  WS:WideString;
begin
  FLjnMimeDecode.DecodeMIME(FMsg.Recipients.EMailAddresses,WS);
  Result:=WS;
end;

function TLjnMailMessage.GetSubject: string;
var
  WS:WideString;
begin
  FLjnMimeDecode.DecodeMIME(FMsg.Subject,WS);
  Result:=Ws;
end;

function TLjnMailMessage.GetTmpPath: string;
begin
  Result:=FTmpPath;
end;

procedure TLjnMailMessage.LoadFromFile(AFileName: string);
begin
  FMsg.Clear;
  FMsg.LoadFromFile(AFileName);
  GetMailContent;
end;

procedure TLjnMailMessage.LoadFromStream(AStream: TStream);
begin
  FMsg.Clear;
  AStream.Position:=0;
  FMsg.LoadFromStream(AStream);
  GetMailContent;
end;

procedure TLjnMailMessage.SetTmpPath(const Value: string);
begin
  FTmpPath:=Value;
  if FTmpPath[Length(FTmpPath)]<>'\' then FTmpPath:=FTmpPath+'\';
end;

procedure TLjnMailMessage.DoAfterMailDecode;
begin
  if Assigned(FAfterMailDecode) then
  begin
    FAfterMailDecode(FMailTextList,FAttachParamList);
  end;
end;

{ TAttachParamList }

procedure TAttachParamList.Add(Attach: TAttachParam);
begin
  FList.Add(Attach);
end;

procedure TAttachParamList.Clear;
var
  i:Integer;
  Obj:TObject;
begin
  for i:=FList.Count-1 downto 0 do
  begin
    Obj:=FList.Items[i];
    //Obj.Free; //Delete ��ʱ����Զ��������
    FList.Delete(i);
  end;
end;

constructor TAttachParamList.Create;
begin
  FList:=TObjectList.Create(True);
end;

destructor TAttachParamList.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TAttachParamList.GetCount: Integer;
begin
  Result:=FList.Count;
end;

function TAttachParamList.GetItem(Index: Integer): TAttachParam;
begin
  Result:=TAttachParam(FList.Items[Index]);
end;

procedure TAttachParamList.SetItem(Index: Integer;
  const Value: TAttachParam);
begin
  FList.Items[Index]:=Value;
end;

{ TMailTextList }

procedure TMailTextList.Add(MailText: TMailText);
begin
  FList.Add(MailText);
end;

procedure TMailTextList.Clear;
var
  i:Integer;
  Obj:TObject;
begin
  for i:=FList.Count-1 downto 0 do
  begin
    Obj:=FList.Items[i];
    //Obj.Free;
    FList.Delete(i);
  end;
end;

constructor TMailTextList.Create;
begin
  FList:=TObjectList.Create(True);
end;

destructor TMailTextList.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TMailTextList.GetCount: Integer;
begin
  Result:=FList.Count;
end;

function TMailTextList.GetHtmlText: Widestring;
var
  i:Integer;
  MailText:TMailText;
  S:WideString;
begin
  S:='';
  for i:=0 to Pred(self.Count) do
  begin
    MailText:=self.Items[i];
    if POS('HTML', UpperCase(MailText.ContentType))>0 then
    begin
      S:=S+MailText.FText;
    end;
  end;
  Result:=S;
end;

function TMailTextList.GetIsHTML: Boolean;
var
  i:Integer;
  MailText:TMailText;
begin
  Result:=False;
  for i:=0 to Pred(self.Count) do
  begin
    MailText:=self.Items[i];
    if POS('HTML', UpperCase(MailText.ContentType))>0 then
    begin
      Result:=True;
      Break;
    end;
  end;
end;

function TMailTextList.GetItem(Index: Integer): TMailText;
begin
  Result:=TMailText(FList.Items[Index]);
end;

function TMailTextList.GetPlainText: WideString;
var
  i:Integer;
  MailText:TMailText;
  S:WideString;
begin
  S:='';
  for i:=0 to Pred(self.Count) do
  begin
    MailText:=self.Items[i];
    if POS('PLAIN', UpperCase(MailText.ContentType))>0 then
    begin
      S:=S+MailText.FText;
    end;
  end;
  Result:=S;
end;

procedure TMailTextList.SetItem(Index: Integer; const Value: TMailText);
begin
  FList.Items[Index]:=Value;
end;

end.
