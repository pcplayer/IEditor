unit MIMEEncodePic;

{-------------------------------------------------------------------------------
  使用浏览器来做编辑器，在浏览器里插入的本地图片或其它的诸如声音等本地媒体文件，
  在浏览器输出的 html 文本里仅仅是一条链接。

  本单元把链接更改为 MIME 的附件描述，再把本地文件作为 MIME 的附件，然后一起输出
  整个编码后的 MIME。

  接收方接收到本单元输出的 MIME，使用 IdMessage 解码，如果有图片附件则也解码出来。

  发送方：
  1. HTML 编写，
  2. 获取 HTML 代码
  3. 从代码里分析出图片，如果图片是本地文件，则对 HTML 代码进行修改，将图片名换成 MIME 的名字；（文字分析状态机）
  4. 把 HTML 代码作为 IdMessage 的文本内容，将本地图片作为 IdMessage 的附件；
  5. 从 IdMessage 取得没有编码的包括附件在内的 MIME 文本。

  接收方：
  1. 接收到 MIME 文本，将文本送进 IdMessage 解码；
  2. 如果有图片附件，解码出来将图片保存为本地文件，并将文本里对 MIME 附件的引用替换为对文件的引用；
  3. 显示：将 IdMessage 解码的文本丢给浏览器显示。

  by pcplayer 2006-3-31

    今天已经升级为支持 delphi 2010 的控件，主要是因为 Indy 版本改了
    现在这里的代码支持 Indy 10，则不支持 Indy 9 了。不兼容。
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
    property ContentTransferEncoding: string read GetContentTransfer write SetContentTransfer; //可能值：quoted-printable or base64 字符串
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
  FMessage.ContentTransferEncoding := '';  // 为空，则默认是 base64 'quoted-printable';
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
  这部分代码，把来自浏览器作为编辑器的 html 代码里的本地图片变为 MIME 的附件
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

   要用状态机找出图片：IMG SRC="" 部分，然后，把图片 D:\center.jpg 提取出来，
   变换成：cid:center.jpg
   然后用：idAttach := TidAttachment.Create(p, Attach);
   其中，p: TidMessageParts; Attach='D:\center.jpg',这样IdMessage将自动把文件的内容导入并编码.
   再将前面变换过图片名为cid 的正文赋予这个部分的正文内容 ，等等。
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
           Status:=mtsIdle; //出错了！
         end
         else AFileName:=AFileName+S[i];
       end;

       mtsQuto2:
       begin
         //走到这里，拿到了整个文件名。但可能要继续循环，拿第二个文件名
         //开始处理拿到的文件名
         TMPFileName := UpperCase(AFileName);

         //AFileName:=UpperCase(AFileName);
         if Pos('HTTP://',TMPFileName)=1 then
         begin
           //http的文件就不替换
           Status:=mtsIdle;
           AFileName:='';
         end
         else
         begin
           //判断扩展名是否是图片文件
           if Length(ExtractFileExt(AFileName))<>3 then Status:=mtsIdle;

           if POS(UpperCase(ExtractFileExt(AFileName)),'JPG.GIF.BMP')=0 then
           begin
             Status:=mtsIdle;
           end;

           if Pos('FILE:///', TMPFileName) = 1 then
           begin
             Delete(AFileName, 1, 8); //去掉 file:///D:/abc/aaa.jpg 的头部。2007-8-21
           end;

           AFileName := StringReplace(AFileName, '/', '\', [rfReplaceAll]); //add by pcplayer 把斜杠全部替换，否则下面的 ExtractFileName 不起作用。2007-8-21

           BFileName:='cid:'+ ExtractFileName(AFileName);
           Delete(S,SrcBegin+1,i-SrcBegin-2); //删除原来的文件名
           Insert(BFileName,S,SrcBegin+1); //在这个位置插入改为 cid:fileName 的串

           //文本替换完了，消息对象里要加上该文件的部分。或者可以把文件名放进一个stringList，最后来装？
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
       idText1.CharSet := Self.GetCurrentCharSet; // 'gb2312'; //todo: 这个应该根据当前输入的字符编码，不该写死
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
       idText1.CharSet := Self.GetCurrentCharSet; // 'gb2312'; //todo: 这个应该根据当前输入的字符编码，不该写死
       idText1.ContentTransfer :=  FMessage.ContentTransferEncoding; //'quoted-printable'; //被 Help 误导了，以为要加双引号！
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
