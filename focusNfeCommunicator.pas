unit focusNfeCommunicator;

interface

uses

  Forms, iniFiles, Classes, sysUtils, Dialogs, windows, IdHTTP;

type

TTipoLog = (tlErro, tlAviso, tlEvento);
TFocusNFeCommunicator = class
  private
    FsendDir: string;
    FToken: string;
    FreceiveDir: string;
    FUrl: string;
    FlogDir: string;
    ref: string; //Feio isso ser uma global, mas por hora acho que � o suficiente
    function getNomeArquivoIni: string;
    procedure ensureDir(nomeDir: string);
    procedure ensureDirs;
    procedure log(tipoLog: TTipoLog; msg: string);
    procedure parseParams;
    procedure criaArquivoPendenciaRetorno(ref: string);
    function getPendenciaRetornoDir: string;
    function getSendDir: string;
    function enviarArquivo(nomeArquivo: string): boolean;
    function getArquivosProcessadosDir: string;
    procedure consultarRetorno(ref: string);
    function getReceiveDir: string;
    procedure baixarArquivos(chave: string);
    function cancelarNota(ref: string): boolean;
    function emitirCartaCorrecao(ref: string): boolean;
  public
    class procedure startProcess;
    constructor create;
  private
    HTTPClient: TIdHTTP;
    property sendDir: string read getSendDir write FsendDir;
    property logDir: string read FlogDir write FlogDir;
    property receiveDir: string read getReceiveDir write FreceiveDir;
    property token: string read FToken write FToken;
    property url: string read FUrl write FUrl;
    procedure readConfiguration;
    procedure processFilesToSend;
    procedure processFilesToReceive;
    function getBaseDir: string;
end;

implementation

uses StrUtils;


class procedure TFocusNFeCommunicator.startProcess;
var
  inicio, fim1, fim2: cardinal;
begin
  with TFocusNFeCommunicator.create do
  begin
    try
      readConfiguration;
      parseParams;
      if ref <> '' then
        criaArquivoPendenciaRetorno(ref);
      Write('Iniciou em ');
      Writeln(GetTickCount);
      processFilesToSend;
      Write('Enviados em ');
      Writeln(GetTickCount);
      processFilesToReceive;
      Write('Retornados em ');
      Writeln(GetTickCount);
    except
      on e: exception do
        log(tlErro, e.Message);
    end;
  end;
end;

constructor TFocusNFeCommunicator.create;
begin
  ref := '';

  HTTPClient := TIdHTTP.Create(nil);
  HTTPClient.HandleRedirects := true;
end;

procedure TFocusNFeCommunicator.readConfiguration;
var
  ini: TIniFile;
  iniFileName: string;
begin
  iniFileName := getBaseDir + getNomeArquivoIni;
  Ini := TIniFile.Create(iniFileName);
  try
    sendDir := Ini.ReadString('Diretorios', 'envio', getBaseDir + 'envios');
    receiveDir := Ini.ReadString('Diretorios', 'retorno', getBaseDir + 'retornos');
    logDir := Ini.ReadString('Diretorios', 'logs', getBaseDir + 'logs');
    url := Ini.ReadString('Conexao', 'url', 'http://producao.acrasnfe.acras.com.br/');
    token := Ini.ReadString('Conexao', 'token', '{token-enviado-pelo-suporte-focusnfe}');
    if not(fileExists(iniFileName)) then
    begin
      Ini.WriteString('Diretorios', 'envio', sendDir);
      Ini.WriteString('Diretorios', 'retorno', receiveDir);
      Ini.WriteString('Diretorios', 'logs', logDir);
      Ini.WriteString('Conexao', 'url', url);
      Ini.WriteString('Conexao', 'token', token);
    end;
    ensureDirs;
  finally
    FreeAndNil(ini);
  end;
end;

procedure TFocusNFeCommunicator.processFilesToSend;
var
  ref: string;
  FindResult: integer;
  SearchRec : TSearchRec;
  nomeArquivo, nomeArquivoDestino: string;
begin
  FindResult := FindFirst(sendDir + '*.nfe', faAnyFile - faDirectory, SearchRec);
  while FindResult = 0 do
  begin
    nomeArquivo := sendDir + SearchRec.Name;
    log(tlEvento, 'Processando arquivo ' + quotedStr(nomeArquivo));
    if enviarArquivo(nomeArquivo) then
    begin
      nomeArquivoDestino := getArquivosProcessadosDir + ExtractFileName(nomeArquivo);
      moveFile(Pchar(nomeArquivo), PChar(nomeArquivoDestino));
      criaArquivoPendenciaRetorno(copy(ExtractFileName(nomeArquivo), 1, length(ExtractFileName(nomeArquivo))- 4));
    end;
    FindResult := FindNext(SearchRec);
  end;

  FindResult := FindFirst(sendDir + '*.can', faAnyFile - faDirectory, SearchRec);
  while FindResult = 0 do
  begin
    nomeArquivo := sendDir + SearchRec.Name;
    ref := copy(SearchRec.Name, 1, length(SearchRec.Name) - 4);
    log(tlEvento, 'Processando arquivo ' + quotedStr(nomeArquivo));
    if cancelarNota(ref) then
    begin
      nomeArquivoDestino := getArquivosProcessadosDir + ExtractFileName(nomeArquivo);
      if not moveFile(Pchar(nomeArquivo), PChar(nomeArquivoDestino)) then
        log(tlErro, 'Imposs�vel mover arquivo para processados: ' + nomeArquivo);
      criaArquivoPendenciaRetorno(ref);
    end;

    FindResult := FindNext(SearchRec);
  end;

  FindResult := FindFirst(sendDir + '*.cce', faAnyFile - faDirectory, SearchRec);
  while FindResult = 0 do
  begin
    nomeArquivo := sendDir + SearchRec.Name;
    ref := copy(SearchRec.Name, 1, length(SearchRec.Name) - 4);
    log(tlEvento, 'Processando arquivo ' + quotedStr(nomeArquivo));
    if emitirCartaCorrecao(ref) then
    begin
      nomeArquivoDestino := getArquivosProcessadosDir + ExtractFileName(nomeArquivo);
      if not moveFile(Pchar(nomeArquivo), PChar(nomeArquivoDestino)) then
        log(tlErro, 'Imposs�vel mover arquivo para processados: ' + nomeArquivo);
      //criaArquivoPendenciaRetorno(ref);
    end;

    FindResult := FindNext(SearchRec);
  end;

end;

procedure TFocusNFeCommunicator.processFilesToReceive;
var
  FindResult: integer;
  SearchRec : TSearchRec;
  nomeArquivo: string;
begin
  FindResult := FindFirst(getPendenciaRetornoDir + '*.ref', faAnyFile - faDirectory, SearchRec);
  while FindResult = 0 do
  begin
    nomeArquivo := getPendenciaRetornoDir + SearchRec.Name;
    log(tlEvento, 'Consultando retorno ' + quotedStr(nomeArquivo));
    consultarRetorno(copy(SearchRec.Name, 0, length(SearchRec.Name) - 4));
    FindResult := FindNext(SearchRec);
  end;
end;

function TFocusNFeCommunicator.getBaseDir: string;
begin
  result := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
end;

function TFocusNFeCommunicator.getPendenciaRetornoDir: string;
begin
  result := IncludeTrailingPathDelimiter(getBaseDir + 'pendenciasRetornos');
end;

function TFocusNFeCommunicator.getArquivosProcessadosDir: string;
begin
  result := IncludeTrailingPathDelimiter(getBaseDir + 'arquivosProcessados');
end;


function TFocusNFeCommunicator.getNomeArquivoIni: string;
begin
  result := 'focusFin.ini';
end;

procedure TFocusNFeCommunicator.ensureDir(nomeDir: string);
begin
  if not(DirectoryExists(nomeDir)) then
    CreateDir(nomeDir);
end;

procedure TFocusNFeCommunicator.ensureDirs;
begin
  ensureDir(sendDir);
  ensureDir(receiveDir);
  ensureDir(receiveDir + 'DANFEs');
  ensureDir(receiveDir + 'XMLs');
  ensureDir(logDir);
  ensureDir(getPendenciaRetornoDir);
  ensureDir(getArquivosProcessadosDir);
end;

procedure TFocusNFeCommunicator.log(tipoLog: TTipoLog; msg: string);
var
  logFile: TStringList;
  nomeLogFile, logMessage: string;
begin
  nomeLogFile := IncludeTrailingPathDelimiter(logDir) +
    formatDateTime('yyyymmdd', date) + '.log';
  logFile := TStringList.create;
  try
    if FileExists(nomeLogFile) then
      logFile.loadFromFile(nomeLogFile);
    logMessage := FormatDateTime('[hh:nn:ss - ', now);
    case tipoLog of
      tlErro: logMessage := logMessage + 'ERRO] ';
      tlAviso: logMessage := logMessage + 'AVISO] ';
      tlEvento: logMessage := logMessage + 'EVENTO] ';
    end;
    logMessage := logMessage + msg;
    logFile.Add(logMessage);
    logFile.SaveToFile(nomeLogFile);
  finally
    FreeAndNil(logFile);
  end;
end;

procedure TFocusNFeCommunicator.parseParams;
var
  i: integer;
  nomeParam, valorParam: string;
begin
  valorParam := '';
  for i := 1 to ParamCount do
  begin
    nomeParam := copy(ParamStr(i), 1, pos('=', ParamStr(i))-1);
    if UpperCase(nomeParam) = 'REF' then
      ref := copy(ParamStr(i), pos('=', ParamStr(i)) + 1, MaxInt)
    else
      log(tlErro, 'Chamado com par�metros inv�lidos: ' + ParamStr(i));
  end;
end;

procedure TFocusNFeCommunicator.criaArquivoPendenciaRetorno(ref: string);
begin
  with TStringList.create do
  begin
    SaveToFile(getPendenciaRetornoDir + ref + '.ref');
    free;
  end;
end;

function TFocusNFeCommunicator.getSendDir: string;
begin
  Result := IncludeTrailingPathDelimiter(FsendDir);
end;

function TFocusNFeCommunicator.getReceiveDir: string;
begin
  Result := IncludeTrailingPathDelimiter(FreceiveDir);
end;

function TFocusNFeCommunicator.enviarArquivo(nomeArquivo: string): boolean;
var
  urlReq: string;
  request, response: TStringStream;
  dados: TStringList;
  refNota: string;
begin
  result := true;
  refNota := copy(extractFileName(nomeArquivo), 1, length(extractFileName(nomeArquivo)) - 4);
  urlReq := '?token=' + token + '&ref=' + refNota;
  response := TStringStream.Create('');
  dados := TStringList.create;
  try
    dados.LoadFromFile(nomeArquivo);
    request := TStringStream.Create(UTF8Encode(dados.Text));
  finally
    FreeAndNil(dados);
  end;
  try
    HTTPClient.Post(url + '/nfe2/autorizar' + urlReq, request, response);
    log(tlEvento, 'Nota enviada com sucesso, ref = ' + refNota);
  except
    on e: EIdHTTPProtocolException do
    begin
      result := false;
      log(tlErro, e.ErrorMessage);
    end;
    on e: Exception do
    begin
      result := false;
      log(tlErro, e.Message);
    end;
  end;
end;

procedure TFocusNFeCommunicator.consultarRetorno(ref: string);
var
  urlReq, chave, valor, chaveNfe, mensagemStatus, mensagemSefaz: string;
  res: TStringList;
  i, separatorPosition: integer;
  final: boolean;
begin
  //final ser� a flag indicando que a resposta � final, autorizado ou n�o autorizado, com isso
  //  podemos apagar o arquivo da pend�ncia de consulta
  final := false;

  urlReq := url + '/nfe2/consultar?token=' + token + '&ref=' + ref;
  res := TStringList.Create;
  try
    res.Text := Utf8ToAnsi(HTTPClient.Get(urlReq));
    res.SaveToFile(receiveDir + ref + '.ret');
    log(tlAviso, 'Salvo arquivo de retorno: ' + QuotedStr(receiveDir + ref + '.ret'));
    for i := 1 to res.Count -1 do
    begin
      separatorPosition := pos(':', res[i]);
      chave := Trim(copy(res[i], 1, separatorPosition - 1));
      valor := Trim(copy(res[i], separatorPosition + 1, MaxInt));

      if chave = 'status' then
        mensagemStatus := valor;
      if chave = 'mensagem_sefaz' then
        mensagemSefaz := valor;
      if chave = 'chave_nfe' then
        chaveNfe := valor;
    end;
    if mensagemStatus = 'processando_autorizacao' then
      log(tlAviso, 'A nota ainda est� em processamento: ' + ref);
    if mensagemStatus = 'erro_autorizacao' then
    begin
      log(tlErro, 'A nota n�o foi autorizada: ' + ref + #13#10 +
        #9 + mensagemSefaz);
      final := true;
    end;
    if mensagemStatus = 'autorizado' then
    begin
      log(tlAviso, 'A nota foi autorizada com sucesso: ' + ref);
      baixarArquivos(chaveNfe);
      final := true;
    end;
    if mensagemStatus = 'cancelado' then
    begin
      log(tlAviso, 'O cancelamento da nota foi autorizado: ' + ref);
      final := true;
    end;

    if final then
      DeleteFile(PChar(getPendenciaRetornoDir + ref + '.ref'));
  finally
    FreeAndNil(res);
  end;
end;

procedure TFocusNFeCommunicator.baixarArquivos(chave: string);
var
  fileName: string;
  fileStream: TMemoryStream;
begin
  //DANFE
  fileName := getReceiveDir + 'DANFEs\' + chave + '.pdf';

  fileStream := TMemoryStream.Create;
  try
    HTTPClient.Get(url + '/notas_fiscais/' + trim(chave) + '.pdf', fileStream);
    fileStream.SaveToFile(fileName);
  finally
    fileStream.free;
  end;

  //XML
  fileName := getReceiveDir + 'XMLs\' + chave + '.xml';

  fileStream := TMemoryStream.Create;
  try
    HTTPClient.Get(url + '/notas_fiscais/' + trim(chave) + '.xml', fileStream);
    fileStream.SaveToFile(fileName);
  finally
    fileStream.free;
  end;
end;

function TFocusNFeCommunicator.cancelarNota(ref: string): boolean;
var
  motivo, urlReq: string;
  response: TStringStream;
  arq: TStringList;
begin
  result := false;
  arq := TStringList.Create;
  response := TStringStream.Create('');
  try try
    arq.LoadFromFile(getSendDir + ref + '.can');
    motivo := trim(arq.GetText);
    urlReq := trim(url + '/nfe2/cancelar?token=' + token +
      '&ref=' + ref + '&justificativa=' + AnsiReplaceStr(motivo, ' ', '+'));
    HTTPClient.Get(urlReq, response);
    log(tlAviso, 'Retorno do cancelamento: ' + response.DataString);
    result := true;
  except
    on e: EIdHTTPProtocolException do
      log(tlErro, e.ErrorMessage);
    on e: Exception do
      log(tlErro, e.Message);
  end;
  finally
    FreeAndNil(response);
    freeAndNil(arq);
  end;
end;

function TFocusNFeCommunicator.emitirCartaCorrecao(ref: string): boolean;
var
  motivo, urlReq: string;
  response: TStringStream;
  arq: TStringList;
begin
  result := false;
  arq := TStringList.Create;
  response := TStringStream.Create('');
  try try
    arq.LoadFromFile(getSendDir + ref + '.can');
    motivo := trim(arq.GetText);
    urlReq := trim(url + '/nfe2/carta_correcao?token=' + token +
      '&ref=' + ref + '&justificativa=' + AnsiReplaceStr(motivo, ' ', '+'));
    HTTPClient.Get(urlReq, response);
    log(tlAviso, 'Retorno do cancelamento: ' + response.DataString);
    result := true;
  except
    on e: EIdHTTPProtocolException do
      log(tlErro, e.ErrorMessage);
    on e: Exception do
      log(tlErro, e.Message);
  end;
  finally
    FreeAndNil(response);
    freeAndNil(arq);
  end;
end;


end.
