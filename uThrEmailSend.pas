unit uThrEmailSend;
// Рассылка email сообщений

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, dxmdaset, HTTPSend, WinInet, synacode, ZDataset,  Generics.Collections, uVars;

type MyTypeA = array of array[0..1] of string;


type
TOptionsProxy = record
  ProxyEnabled: boolean;
  ProxyServer: string;
  ProxyPort: integer;
end;

TMailParam = record
  email: string;
  text: string;
end;

TThrEmail = class(TThread)
  private
    { Private declarations }
    FTag: Integer;
    Progress: integer;
    Pretension: Boolean;
    q: TdxMemData;
    qF: TZQuery;
    listMail: TList<TMailParam>;
  protected
    procedure Execute; override;
    procedure DoTerminate(Sender: TObject);
    procedure SyncThr;
    function GetUrl(arr: MyTypeA; url: string; od_id: string): string;
    function parce_xml(xml: string; od_id:string): string;
  public
    maxEmail: integer;
    text_em:string;
    post_url:string;
    to_wait: integer;
    Proxy: TOptionsProxy;
    FStatus: TSyncFilesStatus;
    constructor Create(q2: TdxMemData; qF2: TZQuery; is_pretension: Boolean = false); reintroduce; overload;
    procedure Stop(Sender: TObject);
    property Terminated;
    property Tag: Integer read FTag write FTag;
   end;


implementation

uses uMain;

procedure GetProxyData(var ProxyEnabled: boolean; var ProxyServer: string; var ProxyPort: integer);
var
  ProxyInfo: PInternetProxyInfo;
  Len: LongWord;
  i, j: integer;
begin
  Len := 4096;
  ProxyEnabled := false;
  GetMem(ProxyInfo, Len);
  try
    if InternetQueryOption(nil, INTERNET_OPTION_PROXY, ProxyInfo, Len)
      then
      if ProxyInfo^.dwAccessType = INTERNET_OPEN_TYPE_PROXY then
      begin
        ProxyEnabled := True;
        ProxyServer := String(ProxyInfo^.lpszProxy);
      end
  finally
    FreeMem(ProxyInfo);
  end;

  if ProxyEnabled and (ProxyServer <> '') then
  begin
    i := Pos('http=', ProxyServer);
    if (i > 0) then
    begin
      Delete(ProxyServer, 1, i + 5);
      j := Pos(';', ProxyServer);
      if (j > 0) then
        ProxyServer := Copy(ProxyServer, 1, j - 1);
    end;
    i := Pos(':', ProxyServer);
    if (i > 0) then
    begin
      ProxyPort := StrToIntDef(Copy(ProxyServer, i + 1, Length(ProxyServer) - i), 0);
      ProxyServer := Copy(ProxyServer, 1, i - 1)
    end
  end;
end;

constructor TThrEmail.Create(q2: TdxMemData; qF2: TZQuery; is_pretension: Boolean = false);
var
  row: TMailParam;
begin
  inherited Create(True);
  q:=q2;
  qF:=qF2;
  FStatus:=thStopped;
  Progress:=0;
  Pretension:=is_pretension;
  if is_pretension then  //отдельная обработка претензий (через список, а не TdxMemData) - в потоке происходила очистка ссылки на TdxMemData
  begin
   listMail:=TList<TMailParam>.Create;
   q2.First;
   while not q2.Eof do
   begin
     row.email:= q2.FieldByName('email').AsString;
     row.text:=  q2.FieldByName('text_mail').AsString;
     listMail.Add(row);
     q2.Next;
   end;
  end;
  GetProxyData(Proxy.ProxyEnabled, Proxy.ProxyServer,Proxy.ProxyPort);
  SyncThr;
  Priority := tpNormal;
  FreeOnTerminate := True;
end;

function TThrEmail.parce_xml(xml: string; od_id:string): string;
begin
  if trim(xml)<>'ok' then
   ShowMessage('Ошибка со стороны сервера:'+ Trim(xml))
  else if not Pretension then
    begin
      //Маркируем email как отправленный (запрос к базе)
      qF.Close;
      qF.SQl.text:='select bul.border_docum_u_set(:order_docum_id)';
      qF.Params.ParamByName('order_docum_id').Value:='{'+od_id+'}';
      qF.ExecSQL;
    end;       
  Result := '0';
end;

function TThrEmail.GetUrl(arr: MyTypeA; url: string; od_id: string): string;
var
  stream: TStringStream;
  BodyS: TStringList;
  HTTP: THTTPSend;
  res2: string;
//  err: string;
//  err_comment: string;
  i: integer;
//  res: integer;
  strpos: integer;
//  tov_name: string;
begin
//  res := 1; //Всегда ошибка
  res2 := '';
  HTTP := THTTPSend.Create;
  try
    if Proxy.ProxyEnabled then
    begin
      HTTP.ProxyHost := Proxy.ProxyServer;
      HTTP.ProxyPort := inttostr(Proxy.ProxyPort);
    end;

    http.MimeType := 'application/x-www-form-urlencoded';
    stream := TStringStream.Create();
    stream.WriteString('&pass=LJGHO(66((!!*F13');

    for i := 0 to High(arr) do
    begin
      stream.WriteString('&' + arr[i][0] + '=' + String(EncodeURLElement(AnsiToUtf8(trim(arr[i][1])))))
    end;
    HTTP.Document.LoadFromStream(stream);
    if HTTP.HTTPMethod('POST', url) then
    begin
      BodyS := TStringList.Create;
      BodyS.LoadFromStream(HTTP.Document);
      res2 := BodyS.Text;
      strpos := pos('PHP Warning', res2);
      if strpos = 0 then
        parce_xml(res2, od_id);
    end
  finally
    HTTP.Free;
  end;
  Result := res2;
end;

procedure TThrEmail.Execute;
var
  multiArray: MyTypeA;
  sh2: string;
  old_em: string;
//  text_mail: string;
  file_pdf: string;
  order_docum_id: string;
  rec: TMailParam;
begin
  try
    try
      FStatus := thRun;
      if not Pretension then
      begin
        if q.RecordCount > 0 then
        begin
          q.First;
          old_em := q.FieldByName('email').AsString;
          while not q.Eof do
          begin
            if file_pdf = '' then
            begin
              file_pdf := q.FieldByName('filePDF').AsString;
              order_docum_id := q.FieldByName('order_docum_id').AsString;
            end
            else
            begin
              file_pdf := file_pdf + ',' + q.FieldByName('filePDF').AsString;
              order_docum_id := order_docum_id + ',' + q.FieldByName('order_docum_id').AsString;
            end;
            if old_em <> q.FieldByName('email').AsString then
            begin
              old_em := q.FieldByName('email').AsString;
              //отправка письма
              SetLength(multiArray, 3);
              multiArray[0][0] := 'email';
              multiArray[0][1] := q.FieldByName('email').AsString;
              sh2 := StringReplace(text_em, '<client_name>', q.FieldByName('fio').AsString, [rfReplaceAll, rfIgnoreCase]);
              multiArray[1][0] := 'txt_e';
              multiArray[1][1] := sh2; //шаблон письма
              multiArray[2][0] := 'file_pdf';
              multiArray[2][1] := file_pdf; //Массив файлов
              GetUrl(multiArray, post_url, order_docum_id);
              sleep(to_wait);
              file_pdf := '';
              order_docum_id := '';
            end;
            inc(Progress);
            Synchronize(SyncThr);
            q.Next;
          end;
          if file_pdf <> '' then
          begin
            SetLength(multiArray, 3);
            multiArray[0][0] := 'email';
            multiArray[0][1] := q.FieldByName('email').AsString;
            sh2 := StringReplace(text_em, '<client_name>', q.FieldByName('fio').AsString, [rfReplaceAll, rfIgnoreCase]);
            multiArray[1][0] := 'txt_e';
            multiArray[1][1] := sh2; //шаблон письма
            multiArray[2][0] := 'file_pdf';
            multiArray[2][1] := file_pdf; //Массив файлов
            GetUrl(multiArray, post_url, order_docum_id);
            sleep(to_wait);
            file_pdf := '';
            order_docum_id := '';
          end;
        end;
        DoTerminate(nil);
      end
      else   //отправка только текста на адрес (из модуля Претензий)
      begin
        SetLength(multiArray, 3);
        for rec in listMail do
        begin
          multiArray[0][0] := 'email';
          multiArray[0][1] := rec.email;
          multiArray[1][0] := 'txt_e';
          multiArray[1][1] := rec.text;
          multiArray[2][0] := 'file_pdf';
          multiArray[2][1] := ''; //пусто, так как нет вложений
          GetUrl(multiArray, post_url, '');
          sleep(to_wait);
          inc(Progress);
          Synchronize(SyncThr);
        end;
        DoTerminate(nil);
      end;
    except
      on E: Exception do
        ShowMessage(e.ClassName + ' ошибка, с сообщением : ' + e.Message);
    end;
  finally  // Помечаем поток как требующий закрытия
    FStatus := thStopped;
    Terminate;
  end;
end;

procedure TThrEmail.SyncThr;  //Отображение данных по работе потока
begin
  fmMain.Statbar.Panels[1].Text:='Запущена рассылка: '+inttostr(Progress)+' из '+inttostr(maxEmail);
end;

procedure TThrEmail.Stop(Sender: TObject);
begin
  if FStatus=thStopped then exit;
  Synchronize(SyncThr);
  DoTerminate(Sender);
  FStatus:=thPause;
  if not Suspended then
  Suspend;
end;

procedure TThrEmail.DoTerminate(Sender: TObject);
begin
  fmMain.Statbar.Panels[1].Text := 'Рассылка завершена отправлено ' + inttostr(Progress) + ' писем из ' + inttostr(maxEmail);
  if Pretension then
    ShowMessage('Рассылка завершена отправлено ' + inttostr(Progress) + ' писем из ' + inttostr(maxEmail));
  FreeAndNil(listMail);
end;


//initialization
 // FlagHasUpdates := False;


end.
