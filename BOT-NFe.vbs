' -- inicializando objetos e constantes para todo o processo -- '
set objFSO = createObject("Scripting.FileSystemObject")


CONST forReading = 1
CONST adTypeBinary = 1
CONST adSaveCreateOverWrite = 2
CONST ForAppending = 8
CONST asASCII = 0
CONST SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056

'ON ERROR RESUME NEXT
' *** *** *** *** *** *** *** 

' -- nome do arquivo de log: AAAAMMDDHHMMSS.log -- '
nomeLog = year(Date) 				   & right("00" & month(Date), 2) 	& right("00" & day(Date), 2) 	 & _
		  right("00" & hour(now()), 2) & right("00" & minute(now()), 2) & right("00" & second(now()), 2) & ".log"

		  
' -- ler arquivo ini -- '
set dictonary = readIniFile("config.ini")


' -- caso o endere�o do log n�o indicado no arquivo INI. Ser� gravado um log tempor�rio na raiz do VBS -- '
if (trim(dictonary("path_log")) = "") then
	' -- cria log temporario -- '
	set tempLogFile = objFSO.createTextFile( nomeLog , true, true )
	
	dim msgErro : msgErro = "[ERRO] Caminho dos arquivos de LOG n�o foi indicado no arquivo INI"
	tempLogFile.writeLine Hour(Now) & ":" & Right("00" & Minute(Now), 2) & ":" & Right("00" & Second(Now), 2) & " -> " & msgErro
	
	tempLogFile.close()
	set tempLogFile = nothing
	
	wscript.echo(msgErro)
	wscript.quit()
end if

' *** *** *** *** *** *** *** 

' -- verifica se a pasta de log est� criada -- '
if ( not( objFSO.folderExists(dictonary("path_log")) ) ) then
	objFSO.createFolder dictonary("path_log")
	'
	checkProcessError("Ocorreu um erro ao criar a pasta de log: " & dictonary("path_log"))
end if

' -- verifica se a pasta de destino est� criada -- '
if ( not( objFSO.folderExists(dictonary("path_destino")) ) ) then
	objFSO.createFolder dictonary("path_destino")
	'
	checkProcessError("Ocorreu um erro ao criar a pasta de destino: " & dictonary("path_log"))
end if


' *** *** *** *** *** *** *** 
' -- cria arquivo de log AAAAMMDDHHMMSS.log -- '
set logFile = objFSO.createTextFile( dictonary("path_log") & "\" & nomeLog , true, true )

checkProcessError("Ocorreu um erro ao criar o arquivo de LOG: " & dictonary("path_log") & "\" & nomeLog)

writeLogFile("############################################################")
writeLogFile("IN�CIO DO PROCESSAMENTO: " & replace(nomeLog, ".log", ""))
writeLogFile("############################################################")
writeLogFile("")

' *** *** *** *** *** *** *** 

' -- tratamento das informa��es -- '
' -- verifica se algum dos itens do arquivo INI n�o foi preenchido -- '
for i=0 to dictonary.count-1
	if ( trim(dictonary.items()(i)) = "" ) then
		writeLogFile("[ERRO] O campo " & dictonary.keys()(i) & " n�o est� preenchido no campo INI")
		fimdoProcessamento()
	end if
next


' -- valida a quantidade de tentativas indicadas no arquivo INI -- '
if ( cInt(dictonary("tentativas")) < 0 ) then
	writeLogFile("[ERRO] A quantidade de tentativas indicadas no arquivo INI n�o � v�lida")
	fimdoProcessamento()
end if

' *** *** *** *** *** *** *** 

' RECUPERAR TOKEN DE ACESSO
Dim tokenContent, token
writeLogFile("[DEBUG] Acessando API para recupera��o do token")

' retorna conte�do JSON da p�gina url_token indicada no arquivo INI
tokenContent = getJSONUrl(dictonary("url_token"), "")

writeLogFile("[DEBUG] Recuperando valor do token a partir da resposta")

' recupera o token de acesso indicado no JSON de retorno "session"
token = getValueFromJSON(tokenContent, "session")
token = "{" & token & "}"

if (trim(token) = "") then
	writeLogFile("[ERRO] O token n�o foi recuperado com sucesso a partir da URL " & dictonary("url_token"))
	fimdoProcessamento()
end if

writeLogFile("[DEBUG] Token recuperado para acesso: " & token)
writeLogFile("")

' *** *** *** *** *** *** *** *** *** *** ***

' RECUPERAR LISTAGEM DAS NFE's
Dim allNFEContent, IDNFE
writeLogFile("[DEBUG] Acessando API para recupera��o da lista de notas fiscais")

' invoca url de listagem das notas fiscais com os parametros indicados no arquivo INI
statusNFE = dictonary("status_nfe")
' 65 NFCe 59 SAT 00 NFSe ou ainda -1 para vir tudo
modeloNFE = dictonary("modelo_nfe")
dataCorrente = year(date) & "-" & right("00" & month(date), 2) & "-" & right("00" & day(date), 2)

' preenche com os parametros necess�rio para a chamada da URL
url_list_nfe = dictonary("url_list_nfe")
url_list_nfe = replace(url_list_nfe, "%CURRENT_DATE%", "=" & dataCorrente)
url_list_nfe = replace(url_list_nfe, "%MODELO_NFE%",   "=" & modeloNFE)
url_list_nfe = replace(url_list_nfe, "%STATUS_NFE%",   "=" & statusNFE)

writeLogFile("[DEBUG] Acesso a API de notas fiscais: " & url_list_nfe)

allNFEContent = getJSONUrl(dictonary("url_list_nfe"), token)

' recupera todas as NFE pelo ID indicado no JSON de retorno IDNFE
allNFEContent = getValueFromJSON(allNFEContent, "IDNFE")

'
writeLogFile("[DEBUG] Iniciando itera��es das notas fiscais...")
allNFEContent = split(allNFEContent, ",")

dim totalNotas
totalNotas = 0

' -- itera��o das notas fiscais
for i=0 to uBound(allNFEContent)
	
	IDNFE = cleanString(allNFEContent(i))
	
	' busca o caminho onde o XML ser� salvo
	destino = getDestino(IDNFE)
	
	' *** *** *** *** *** *** *** *** *** *** ***
	dim tentativas
	' quantidade de tentativas indicadas no arquivo INI
	tentativas = dictonary("tentativas")
	
	do while (tentativas > 0)		
		' verifica se o arquivo existe
		if not(objFSO.fileExists(destino)) then
			writeLogFile("")
			writeLogFile("[DEBUG] Iniciando download da NFE: " & IDNFE)
			
			' salva o arquivo em disco
			'saveFile dictonary("url_download") & "=" & IDNFE, destino, token
			saveFile dictonary("url_download"), destino, token
			
			writeLogFile("[DEBUG] Arquivo " & destino & " salvo com sucesso")
			totalNotas = totalNotas + 1
			wscript.sleep(500)
		end if
		
		tentativas = tentativas-1		
	loop
	
next

writeLogFile("")

if (totalNotas > 0) then
	writeLogFile("[DEBUG] Total de notas recuperadas com sucesso: " & i)
else
	writeLogFile("[DEBUG] Nenhuma nota fiscal encontrada")	
end if

fimdoProcessamento()

' *** *** *** *** *** *** *** *** *** *** ***


' #################################
' fun��o para salvar o arquivo em disco
function saveFile(urlDownload, destino, token)
	Dim xmlDoc:  set xmlDoc  = createobject("MSXML2.ServerXMLHTTP")
	Dim oStream: set oStream = createobject("Adodb.Stream")
	
	' chama API de download da NFE
	xmlDoc.open "GET", urlDownload, false
	xmlDoc.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
	xmlDoc.setRequestHeader "WTS-Session", token
	xmlDoc.send()
	
	checkProcessError("Ocorreu um erro ao realizar o download do arquivo em " & urlDownload)

	with oStream
		.type = 1 '//bin�rio
		.open()
		.write xmlDoc.responseBody
		.saveToFile destino, 1
	end with
	
	checkProcessError("Ocorreu um erro ao salvar o arquivo " & destino)
	
	oStream.close()
	set oStream = nothing
	set xmlDoc  = nothing
end function


' -- fun��o para escrever no arquivo de log -- '
function writeLogFile(byVal logText)
	logFile.writeLine Hour(Now) & ":" & Right("00" & Minute(Now), 2) & ":" & Right("00" & Second(Now), 2) & " -> " & logText
end function


' -- fun�ao para escrever no arquivo de log e finalizar o processamento -- '
function fimdoProcessamento()
	' ------------------------------------- '
	writeLogFile("")
	writeLogFile("############################################################")
	writeLogFile("FIM DO PROCESSAMENTO: " & year(Date) 				    & right("00" & month(Date), 2) 	 & right("00" & day(Date), 2) 	& _
										   right("00" & hour(now()), 2) & right("00" & minute(now()), 2) & right("00" & second(now()), 2) )
	writeLogFile("############################################################")
	' ------------------------------------- '
		
	logFile.close()
	set logFile = nothing
	set objFSO = nothing

	wscript.quit()
end function


' -- fun��o para  ler o arquivo INI e popular em um objeto do tipo dictonary
' -- Importante: o arquivo INI deve estar na mesma pasta do arquivo VBS
function readIniFile(nomeArquivo)
	set dictonary  = createObject("Scripting.Dictionary")
	
	set arquivoIni = objFSO.openTextFile(nomeArquivo)
	
	' -- caso ocorra um leitura do arquivo INI, ser� criado um log temporario, pois nesse momento n�o foi ainda lido a pasta de destino dos LOG's
	if (err <> 0) then
		' -- cria log temporario -- '
		set tempLogFile = objFSO.createTextFile( nomeLog , true, true )
	
		msgErro = "Ocorreu um erro ao ler o arquivo " & nomeArquivo & ". Erro: " & err.description
		tempLogFile.writeLine Hour(Now) & ":" & Right("00" & Minute(Now), 2) & ":" & Right("00" & Second(Now), 2) & " -> " & msgErro
	
		wscript.echo(msgErro)
		wscript.quit()
	end if
	
	dim linha, chave, valor
  
	do until (arquivoIni.atEndOfStream)
		linha = trim( arquivoIni.readLine() )
	 
		' -- despreza linha que iniciam on "[" -- '
		if not( left(linha, 1) = "[" ) then
			if ( trim(linha) <> "" ) then			
				valor = split(linha, "=")
				' -- exemplo: username=user
				' -- key: username | value: user
				dictonary.add trim(valor(0)), trim(valor(1))
			end if
        end if     
	loop
	
	arquivoIni.close()
	set arquivoIni = nothing
  
	set readIniFile = dictonary
end function


' -- fun��o para verificar se ocorreu algum erro no processamento e, caso positivo, exibir no log da aplica��o -- '
function checkProcessError(mensagem)
	if (err <> 0) then
		writeLogFile("[ERRO] " & mensagem)
		writeLogFile("[ERRO] " & err.description)
		err.clear()
		fimdoProcessamento()
	end if
end function


' -- fun��o para recuperar uma p�gina com retorno JSON. Utilizada para recuperar o token de acesso ou listagem das NFE -- '
function getJSONUrl(url, token)
	set objHTTP = createObject("WinHttp.WinHttpRequest.5.1")

	objHTTP.open "GET", url, false
	
	' recuperar o token ou acessar a lista de NFE?
	if (trim(token) = "") then
		objHTTP.setRequestHeader "WTS-Authorization", dictonary("wts_authorization")
		objHTTP.setCredentials dictonary("username"), dictonary("password"), 0
	else
		objHTTP.setRequestHeader "WTS-Session", token
	end if
	
	objHTTP.send()
	checkProcessError("Ocorreu um erro ao acessar a url " & url)

	' -- recupera o stream do objeto -- '
	for i=1 to lenB( objHTTP.responseBody() )
		content = content & Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
	next

	checkProcessError("Ocorreu um erro ao recuperar o conte�do da url" & url)

	set objHTTP = nothing
	getJSONUrl = content
end function


' -- fun��o para recuperar o valor do token, identificado pela chave session, a partir do arquivo JSON -- '
function getValueFromJSON(tokenStr, keyValueID)
	set objRegExp 	  = new RegExp
	objRegExp.global  = true
	
	' -- por expressao regular, cria os separadores do arquivo json: no caso da API do token h� apenas 1 -- '
	objRegExp.pattern = "[\[\]\{\}""]+"
	
	tokenStr = replace(tokenStr, "},{", "||")
	tokenStr = objRegExp.replace(tokenStr, "")
	tokenStr = replace( replace(tokenStr, "[", ""), "]", "")
	arrJSON  = split(tokenStr,"||")

	strReturn = ""

	for each reg in arrJSON
	
		' -- os campos do json s�o separados por "," -- '
		JSONLine = split(reg, ",")
		
		for i=0 to uBound(JSONLine)
		
			' -- os campos de um registro s�o separados por ":" -- '
			keyValue = split(JSONLine(i), ":")
		
			' -- recupera o valor da key do json. Exemplo: session:{8306AC3E-CEF643F6-8137-A5AAFA69F3CD}
			' -- key=session | key={8306AC3E-CEF643F6-8137-A5AAFA69F3CD}
			key = keyValue(0)
			key = cleanString( replace(key,"'","") )

			' -- verifica se a key � a session (que possui o token) -- '
			if ( inStr(uCase(key), uCase(keyValueID)) > 0 ) then
				value = trim(keyValue(1))
				value = replace(value,"'","")
				
				strReturn = cleanString(strReturn) & value & ","
			end if
			
		next
	next
	
	checkProcessError("Ocorreu um erro ao formatar o arquivo JSON " & tokenStr)
	
	set objRegExp = nothing
	
	' retira virgula
	if (strReturn <> "") then
		strReturn = left(strReturn, len(strReturn)-1)
	end if
	
	getValueFromJSON = strReturn
	
end function


' -- fun��o para recuperar o caminho destino onde a nota fiscal ser� salva -- '
function getDestino(fileName)

	dim slash, path_destino
	path_destino = dictonary("path_destino")
	
	' verifica se o caminho est� em um servidor windows ou unix
	if ( inStr(path_destino, "/") ) then
		slash = "/"
	else
		slash = "\"
	end if
	
	if ( right(path_destino, 1) = slash ) then
		path_destino = path_destino & fileName
	else
		path_destino = path_destino & slash & fileName
	end if
	
	getDestino = path_destino

end function


' -- fun��o para remover caracteres do arquivo JSON -- '
function cleanString(str)
    str = replace(str, vbTab,  "")
	str = replace(str, vbCrLf, "")
	str = replace(str, chr(10), "")
	str = replace(str, chr(13), "")

    do while (inStr(str, " "))
        str = replace(str, " ", "")
    loop

    cleanString = trim(str)
end function