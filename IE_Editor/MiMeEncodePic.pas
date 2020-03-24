unit MIMEEncodePic;

{-------------------------------------------------------------------------------
  ʹ������������༭����������������ı���ͼƬ�����������������ȱ���ý���ļ���
  ������������ html �ı��������һ�����ӡ�

  ����Ԫ�����Ӹ���Ϊ MIME �ĸ����������ٰѱ����ļ���Ϊ MIME �ĸ�����Ȼ��һ�����
  ���������� MIME��

  ���շ����յ�����Ԫ����� MIME��ʹ�� IdMessage ���룬�����ͼƬ������Ҳ���������

  ���ͷ���
  1. HTML ��д��
  2. ��ȡ HTML ����
  3. �Ӵ����������ͼƬ�����ͼƬ�Ǳ����ļ������ HTML ��������޸ģ���ͼƬ������ MIME �����֣������ַ���״̬����
  4. �� HTML ������Ϊ IdMessage ���ı����ݣ�������ͼƬ��Ϊ IdMessage �ĸ�����
  5. �� IdMessage ȡ��û�б���İ����������ڵ� MIME �ı���

  ���շ���
  1. ���յ� MIME �ı������ı��ͽ� IdMessage ���룻
  2. �����ͼƬ���������������ͼƬ����Ϊ�����ļ��������ı���� MIME �����������滻Ϊ���ļ������ã�
  3. ��ʾ���� IdMessage ������ı������������ʾ��

  by pcplayer 2006-3-31

    �����Ѿ�����Ϊ֧�� delphi 2010 �Ŀؼ�����Ҫ����Ϊ Indy �汾����
    ��������Ĵ���֧�� Indy 10����֧�� Indy 9 �ˡ������ݡ�
    pcplayer 2010-5-31
--------------------------------------------------------------------------------}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, IdBaseComponent, IdMessage,
  IdText, IdMessageParts, IdAttachment, IdAttachmentFile, JclCharSets;

type

  TLjnMessageEncodePic = class
  private
    FMessage: TIdMessage;

    function GetCurrentCharSet: string;

    function GetContentTransfer: string;
    procedure SetContentTransfer(const Value: string);
  public
    constructor Create;
    destructor Destroy;override;

    function MessageEmbedPic(AMessage: WideString): WideString;
    property ContentTransferEncoding: string read GetContentTransfer write SetContentTransfer; //����ֵ��quoted-printable or base64 �ַ���
  end;

implementation

type
 TFindStatus=(mtsIdle,mtsLeft,mtsI,mtsM,mtsG,mtsS,mtsR,mtsC,mtsEqu,mtsQuto1,
   mtsQuto2,mtsRight);

{ TLjnMessageEncodePic }

constructor TLjnMessageEncodePic.Create;
begin
  FMessage := TIdMessage.Create(nil);
  FMessage.Clear;
  FMessage.ContentTransferEncoding := '';  // Ϊ�գ���Ĭ���� base64 'quoted-printable';
  FMessage.CharSet := Self.GetCurrentCharSet; // 'gb2312';
end;

destructor TLjnMessageEncodePic.Destroy;
begin
  FMessage.Free;
  inherited;
end;

function TLjnMessageEncodePic.GetContentTransfer: string;
begin
  if Assigned(FMessage) then Result := FMessage.ContentTransferEncoding;
end;

function TLjnMessageEncodePic.GetCurrentCharSet: string;
var
  CP: Word;
begin
  CP := Windows.GetACP;
  Result := JclCharSets.CharsetNameFromCodePage(CP);
end;

function TLjnMessageEncodePic.MessageEmbedPic(
  AMessage: WideString): WideString;
var
  S: WideString;
  i,SrcBegin,SrcEnd:Integer;
  Status:TFindStatus;
  AFileName,BFileName, TMPFileName:string;
  SL,SL1:TStringList;
  p: TidMessageParts;
  idText1,idText2:TIdText;
  idAttach: TidAttachment;
  StrStream: TStringStream;
begin
{--------------------------------------------------------------------------
  �ⲿ�ִ��룬�������������Ϊ�༭���� html ������ı���ͼƬ��Ϊ MIME �ĸ���
----------------------------------------------------------------------------}
  SL:=TStringList.Create;
  S := AMessage;
  if Length(S) = 0 then
  begin
    Result := S;
    Exit;
  end;

  {------------------------------------------------------------------------
   <BODY><P>asdfasdf</P>
   <P>234324</P>
   <P><IMG alt="" hspace=0 src="D:\center.jpg" align=baseline border=0></P>
   <P>12344455 5666</P>
   <P>&nbsp;</P></BODY>

   Ҫ��״̬���ҳ�ͼƬ��IMG SRC="" ���֣�Ȼ�󣬰�ͼƬ D:\center.jpg ��ȡ������
   �任�ɣ�cid:center.jpg
   Ȼ���ã�idAttach := TidAttachment.Create(p, Attach);
   ���У�p: TidMessageParts; Attach='D:\center.jpg',����IdMessage���Զ����ļ������ݵ��벢����.
   �ٽ�ǰ��任��ͼƬ��Ϊcid �����ĸ���������ֵ��������� ���ȵȡ�
 ------------------------------------------------------------------------------}

  try
   Status:=mtsIdle;
   i:=1;
   SL.Clear;
   while True do                
   begin
     case Status of
       mtsIdle: if S[i]='<' then Status:=mtsLeft;

       mtsLeft:
       begin
         if S[i]=' ' then Continue
         else
         if ((S[i]='I') or (S[i]='i')) then Status:=mtsI
         else Status:=mtsIdle;
       end;

       mtsI:
       begin
         if ((S[i]='M') or (S[i]='m')) then Status:=mtsM else Status:=mtsIdle;
       end;

       mtsM:
       begin
         if ((S[i]='G') or (S[i]='g')) then Status:=mtsG else Status:=mtsIdle;
       end;

       mtsG:
       begin
         if S[i]='>' then Status:=mtsIdle
         else
         if ((S[i]='S') or (S[i]='s')) then Status:=mtsS;
       end;

       mtsS:
       begin
         if ((S[i]='R') or (S[i]='r')) then Status:=mtsR else Status:=mtsG;
       end;

       mtsR:
       begin
         if ((S[i]='C') or (S[i]='c')) then Status:=mtsC else Status:=mtsG;
       end;

       mtsC:
       begin
         if (S[i]='=') then Status:=mtsEqu else Status:=mtsG;
       end;

       mtsEqu:
       begin
         if (S[i]='"') then
         begin
           Status:=mtsQuto1;
           AFileName:='';
           SrcBegin:=i;
         end
         else Status:=mtsG;
       end;

       mtsQuto1:
       begin
         if S[i]='"' then
         begin
           Status:=mtsQuto2;
           SrcEnd := i;
         end
         else
         if S[i]='>' then
         begin
           Status:=mtsIdle; //�����ˣ�
         end
         else AFileName:=AFileName+S[i];
       end;

       mtsQuto2:
       begin
         //�ߵ�����õ��������ļ�����������Ҫ����ѭ�����õڶ����ļ���
         //��ʼ�����õ����ļ���
         TMPFileName := UpperCase(AFileName);

         //AFileName:=UpperCase(AFileName);
         if Pos('HTTP://',TMPFileName)=1 then
         begin
           //http���ļ��Ͳ��滻
           Status:=mtsIdle;
           AFileName:='';
         end
         else
         begin
           //�ж���չ���Ƿ���ͼƬ�ļ�
           if Length(ExtractFileExt(AFileName))<>3 then Status:=mtsIdle;

           if POS(UpperCase(ExtractFileExt(AFileName)),'JPG.GIF.BMP')=0 then
           begin
             Status:=mtsIdle;
           end;

           if Pos('FILE:///', TMPFileName) = 1 then
           begin
             Delete(AFileName, 1, 8); //ȥ�� file:///D:/abc/aaa.jpg ��ͷ����2007-8-21
           end;

           AFileName := StringReplace(AFileName, '/', '\', [rfReplaceAll]); //add by pcplayer ��б��ȫ���滻����������� ExtractFileName �������á�2007-8-21

           BFileName:='cid:'+ ExtractFileName(AFileName);
           Delete(S,SrcBegin+1,i-SrcBegin-2); //ɾ��ԭ�����ļ���
           Insert(BFileName,S,SrcBegin+1); //�����λ�ò����Ϊ cid:fileName �Ĵ�

           //�ı��滻���ˣ���Ϣ������Ҫ���ϸ��ļ��Ĳ��֡����߿��԰��ļ����Ž�һ��stringList�������װ��
           SL.Add(AFileName);
           AFileName:='';
           Status:=mtsIdle;
         end;
       end;

       mtsRight:;
     end; //case

     Inc(i);
     if i>Length(S) then Break;
   end;

   FMessage.Recipients.Add.Address:='pcplayer@sina.com';

   //FMessage.Encoding := meMIME; //add by pcplayer for indy 10
   FMessage.Date := Now;
   if SL.Count>0 then
   begin
     p := FMessage.MessageParts;
     FMessage.ContentType := 'multipart/alternative'; // 'multipart/mixed';

     SL1:=TStringList.Create;
     try
       idText1 := TidText.Create(p, nil);
       SL1.Text:=S;
       idText1.CharSet := Self.GetCurrentCharSet; // 'gb2312'; //todo: ���Ӧ�ø��ݵ�ǰ������ַ����룬����д��
       idText1.ContentTransfer := FMessage.ContentTransferEncoding; // 'quoted-printable';
       idText1.Body := SL1;
     finally
       SL1.Free;
     end;

     idText1.ContentType := 'text/html';

     idText2 := TidText.Create(p);
     idText2.ContentType := 'text/plain';
     idText2.Body.Text := 'This is a HTML Message by pcplayer';

     for i:=0 to SL.Count-1 do
     begin
       AFileName:=SL.Strings[i];

       idAttach := TIdAttachmentFile.Create(p, AFileName);//'c:\sm101yellow.jpg');
       idAttach.ContentType := 'image/'+LowerCase(ExtractFileExt(AFileName));
       idAttach.ContentDisposition := 'inline';
       idAttach.ExtraHeaders.Values['content-id'] := ExtractFileName(AFileName);// 'sm101yellow.jpg';
       idAttach.DisplayName := ExtractFileName(AFileName);
     end;
   end
   else
   begin
     SL1:=TStringList.Create;
     try
       p := FMessage.MessageParts;

       SL1.Text:=AMessage;
       idText1 := TidText.Create(p, nil);
       idText1.CharSet := Self.GetCurrentCharSet; // 'gb2312'; //todo: ���Ӧ�ø��ݵ�ǰ������ַ����룬����д��
       idText1.ContentTransfer :=  FMessage.ContentTransferEncoding; //'quoted-printable'; //�� Help ���ˣ���ΪҪ��˫���ţ�
       idText1.ContentType := 'text/html';
       idText1.Body := SL1;
       FMessage.ContentType := 'text/html';
     finally
       SL1.Free;
     end;
   end;

   StrStream := TStringStream.Create('');

   try
     try
       FMessage.SaveToStream(StrStream);
     except
     end;
   finally
     Result := StrStream.DataString;
   end;
 finally
   SL.Free;
 end;
end;

procedure TLjnMessageEncodePic.SetContentTransfer(const Value: string);
begin
  if Assigned(FMessage) then FMessage.ContentTransferEncoding := Value;
end;

end.
