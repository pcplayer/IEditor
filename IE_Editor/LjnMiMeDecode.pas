unit LjnMiMeDecode;
{-----------------------------------------------------------------------------
  我自己写的对 EMail 中的 MIME 进行解码的代码
  pcplayer 2006-1-22
----------------------------------------------------------------------------- }

interface

uses
  Windows, Messages, SysUtils, StrUtils, Variants, Classes,Controls,
  IdCoderQuotedPrintable, IdCoderUUE, IdBaseComponent,IdCoder, IdCoder3to4,
  IdCoderMIME,ComCtrls,mshtml,ActiveX,OleCtrls,SHDocVw,IdMessage;

type
  TFenXiZT=(fxKS,fxDeng,fxBianMa,fxQB,fxQBEnd,fxZW,fxZWEnd);
  
  TLjnMIMEDecode=class
  private
    IdDecoderMIME: TIdDecoderMIME;
    IdDecoderUUE: TIdDecoderUUE;
    IdDecoderQuotedPrintable: TIdDecoderQuotedPrintable;

    function GetCodePageByStr(CodePageStr:string):word;
  public
    constructor Create;
    destructor Destroy;override;

    function DecodeBase64(Base64Str:string):WideString;
    function DecodeQuotedPrintable(StrIn:string):WideString;
    function StringToWideString(const S: string; CodePage: Word): WideString;
    function DecodeMIME(Sin: string; var SOut: WideString): Boolean;
    function DecodeMailBody(ContentTransferEncodingStr,CharSetStr,
      ContentType,BodyText:string;Encoding:TIdMessageEncoding):WideString;
    function ChangePicInHtml(Sin:string):string;
  end;

procedure ShowTextInWebBrowser(WS:WideString; WB:TWebBrowser);

implementation

procedure ShowTextInWebBrowser(WS:WideString; WB:TWebBrowser);
var
  vv: Variant;
  HTMLDocument: IHTMLDocument2;
begin
  vv := VarArrayCreate([0, 0], varVariant);
  HTMLDocument := WB.Document as IHTMLDocument2;

  vv[0] := WS;

  HTMLDocument.write(pSafearray(TVarData(vv).VArray));
  HTMLDocument.close;
end;


{ TLjnMIMEDecode }

function TLjnMIMEDecode.ChangePicInHtml(Sin: string): string;
var
  OutStr:string;
  tmpStr:string;
begin
  OutStr:=Sin;
  tmpStr:=UpperCase(Sin);
  if POS('<IMG',tmpStr)>0 then
  begin
  end;
end;

constructor TLjnMIMEDecode.Create;
begin
  IdDecoderMIME := TIdDecoderMIME.Create(nil);
  IdDecoderUUE := TIdDecoderUUE.Create(nil);
  IdDecoderQuotedPrintable := TIdDecoderQuotedPrintable.Create(nil);
end;

function TLjnMIMEDecode.DecodeBase64(Base64Str: string): WideString;
var
  TS:TStringList;
  S:string;
  i:Integer;
begin
  TS:=TStringList.Create;
  try
    TS.Text:=Base64Str;
    for i:=0 to TS.Count-1 do
    begin
      S:=TS.Strings[i];
      Result:=Result+IdDecoderMIME.DecodeString(S); //  .DecodeToString(S);
    end;
  except
    TS.Free;
  end;
end;

function TLjnMIMEDecode.DecodeMailBody(ContentTransferEncodingStr,
  CharSetStr, ContentType, BodyText: string; Encoding:TIdMessageEncoding): WideString;
var
  CodePage:Word;
begin
{-------------------------------------------------------------------
  EMail 正文解码
------------------------------------------------------------------------------}
  if UpperCase(ContentTransferEncodingStr)='BASE64' then
  begin
    Result:=DecodeBase64(BodyText);
  end
  else
  if UpperCase(ContentTransferEncodingStr) = 'QUOTED-PRINTABLE' then
  begin
    Result:=DecodeQuotedPrintable(BodyText);
  end
  else
    Result:=BodyText;

  if UpperCase(CharSetStr)='GB2312' then
    CodePage:=936
  else
  if UpperCase(CharSetStr)='BIG5' then
    CodePage:=950
  else
  if UpperCase(CharSetStr)='UTF-8' then
    CodePage:=CP_UTF8
  else
    CodePage:=GetACP; //如果实在是不知道的编码,就按它自己的系统的当前编码来做.权宜之计.{ToDo: 有必要多高清楚几种编码}

  //Result:=StringToWideString(Result,CodePage);

  if Pos('TEXT/PLAIN',UpperCase(ContentType))>0 then      //把空白的普通文本变成 HTML 的换行模式，否则在浏览器里显示不会换行。
    Result:=StringReplace(Result,#13,'<br>',[rfReplaceAll]);
end;

function TLjnMIMEDecode.DecodeMIME(Sin: string;
  var SOut: WideString): Boolean;
var
  i,j,k,BIndex,EIndex,KBIndex,KEIndex:Integer;
  SIn1,S,S2,QBStr:string;
  QB:Char;
  ZT:TFenXiZT;
  WS:WideString;
  CodePageStr:string;
  CodePage:Word;
begin
  //用状态机的方式来分析字符串
  CodePageStr:='';
  SIn1:=Trim(Sin);
  Result:=False;
  if Length(SIn1)=0 then Exit;
  if Length(SIn1)=2 then
  begin
    SOut:=SIn1;
    Exit;
  end;

  WS:='';
  ZT:=fxKS;
  k:=Pos('=',SIn1);
  if k=0 then
  begin
    SOut:=Sin1;
    Exit;
  end
  else
  if k>1 then
  begin
    SOut:=Copy(SIn1,1,k-1);
  end;
  //开始逐个字节进行分析

  KBIndex:=0; //如果开始不是 = 开始的编码字符串，而是普通字符串，则这就是普通字符串的开始位置
  KEIndex:=0; //这是普通字符串的结束位置
  for i:=k to Length(SIn1) do
  begin
    case ZT of
      fxKS:
      begin
        if SIn1[i]='=' then
        begin
          ZT:=fxDeng;
          //j:=i; //纪录开始的位置
        end
        else
        if SIn1[i]=' ' then
        begin
          if Length(SOut)>0 then SOut:=SOut+SIn1[i]; //如果是两条编码中间的空格，则把空格加上去。
        end
        else
        begin
          if KBIndex = 0 then KBIndex:=i else KEIndex:=i;
        end;
      end;

      fxDeng:
      begin
        if SIn1[i]='?' then
        begin
          ZT:=fxBianMa;
        end
        else
        begin
          Break;
        end;
      end;

      fxBianMa:
      begin
        if SIn1[i]='?' then ZT:=fxQB
        else
        begin
          CodePageStr:=CodePageStr+SIn1[i];
        end;
      end;

      fxQB:
      begin
        QBStr:=SIn1[i];
        QBStr:=UpperCase(QBStr);
        QB:=QBStr[1];
        ZT:=fxQBEnd;
      end;

      fxQBEnd:
      begin
        if SIn1[i]='?' then
        begin
          ZT:=fxZW;
          BIndex:=i+1; //编码正式开始的位置
        end
        else
        begin
          Break;
        end;
      end;

      fxZW: //正文
      begin
        if SIn1[i]='?' then ZT:=fxZWEnd;
        EIndex:=i-1; //编码结束的位置
      end;

      fxZWEnd: //正文结束
      begin
        if SIn1[i]='=' then
        begin
          //正式取出来一条编码的数据
          S:=Copy(SIn1,BIndex,EIndex-BIndex+1);

          case QB of
            'B':WS:= DecodeBase64(S);
            'Q':WS:=DecodeQuotedPrintable(S);
          end;
          CodePageStr:=UpperCase(CodePageStr);
          CodePage:=GetCodePageByStr(CodePageStr);
          WS:=StringToWideString(WS,CodePage);


          if Length(SOut)=0 then SOut:=WS else SOut:=SOut+WS;
          ZT:=fxKS;
          Result:=True;
        end
        else
        begin
          ZT:=fxZW; //如果紧跟的不是 = 号，则状态退回去。
        end;
      end;
    end; //case
  end;

  if KBIndex>0 then
  begin
    S:=Copy(SIn1,KBIndex,KEIndex-KBIndex+1);
    SOut:=Sout+S;
  end;
  
  if not Result then SOut:=SIn1;
end;

function TLjnMIMEDecode.DecodeQuotedPrintable(StrIn: string): WideString;
var
  SL:TStringList;
  S:string;
  i:Integer;
begin
  SL:=TStringList.Create;
  try
    SL.Text:=StrIn;

    for i:=0 to SL.Count-1 do
    begin
      S:=SL.Strings[i];
      Result:=Result+IdDecoderQuotedPrintable.DecodeString(S); //  .DecodeToString(S);
    end;
    //Result:=IdDecoderQuotedPrintable.DecodeToString(StrIn);
  except
    SL.Free;
  end;
end;

destructor TLjnMIMEDecode.Destroy;
begin
  IdDecoderMIME.Free;
  IdDecoderUUE.Free;
  IdDecoderQuotedPrintable.Free;
  inherited;
end;

function TLjnMIMEDecode.GetCodePageByStr(CodePageStr: string): word;
var
  Cr:string;
begin
  Cr:=UpperCase(CodePageStr);
  if Cr='GB2312' then Result:=936
  else
  if Cr='BIG5' then Result:=950
  else
  if Cr='UTF-8' then Result:=CP_UTF8
  else
  if Cr='ISO-10646' then Result:=1200
  else
    Result:=CP_OEMCP; //其它编码的字符串描述要另外找。
end;

function TLjnMIMEDecode.StringToWideString(const S: string;
  CodePage: Word): WideString;
var
  InputLength, OutputLength: Integer;
begin
  //todo: 如果用于 Delphi 2010，似乎用不着转换了！
  InputLength := Length(S);
  OutputLength := MultiByteToWideChar(CodePage, 0, PAnsiChar(S), InputLength, nil, 0);
  SetLength(Result, OutputLength);
  MultiByteToWideChar(CodePage, 0, PAnsiChar(S), InputLength, PWideChar(Result), OutputLength);
end;

end.
