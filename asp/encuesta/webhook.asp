<%@ Language="VBScript" %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"




Function LeerAccessToken()
    On Error Resume Next
    Dim fso, f, tokenPath, token
    token = ""
	
	Call CopiarToken()

    tokenPath = "C:\Webhook\whatsapp_token.txt"

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tokenPath) Then
        Set f = fso.OpenTextFile(tokenPath, 1, False)
        token = Trim(f.ReadAll)
        f.Close
        Set f = Nothing
    End If
    Set fso = Nothing

    If Len(token) = 0 Then
        token = "TOKEN_NO_ENCONTRADO"
    End If

    LeerAccessToken = token
End Function


Dim method
method = Request.ServerVariables("REQUEST_METHOD")

' CONFIGURACIÓN
Const VERIFY_TOKEN = "MI_TOKEN_SEGURO"  ' mismo valor que pusiste en Meta
Const ACCESS_TOKEN = "EAASuyx8Ki50BPgWZCraKFKN4cRM2sY6paNAB7Wagf5I2c9T4SB3LKhukN6pLBEmgqpJZA3alVlkWydMd7BqLBZCvOCOBzbL1S4GonElAFsjkjICdkgjKtgSonVQZBzFVA8I14BDgWxVQC8boiphr9gcMlhZCBRqBvYu9gc5uUfCkjSeZAQ6DftuKR2MzhiZAZBm3C3petuvNvOY9"           ' <-- tu token válido de Meta (temporal o permanente)
Const PHONE_NUMBER_ID = "556834667519231" ' el Phone Number ID del número productivo

' ==== GET: Verificación inicial ====
If method = "GET" Then
    Dim mode, token, challenge
    mode = Request.QueryString("hub.mode")
    token = Request.QueryString("hub.verify_token")
    challenge = Request.QueryString("hub.challenge")

    If mode = "subscribe" And token = VERIFY_TOKEN Then
        Response.Write challenge
        Response.Status = "200 OK"
    Else
        Response.Status = "403 Forbidden"
    End If

' ==== POST: Mensajes entrantes ====
ElseIf method = "POST" Then
    Dim rawData, strData
    rawData = Request.BinaryRead(Request.TotalBytes)
    strData = BytesToString(rawData)

    ' Registrar log
    Dim fso, logFile
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile(Server.MapPath("/webhook/webhook_log.txt"), 8, True)
    logFile.WriteLine Now() & " - " & strData & LeerAccessToken() & "Usuario IIS actual: " & CreateObject("WScript.Network").UserName
    logFile.Close
    Set logFile = Nothing
    Set fso = Nothing

    ' Procesar JSON entrante
    Dim json, messages, fromNumber, msgBody
    On Error Resume Next
    Set json = ParseJson(strData)

    If Not json Is Nothing Then
        Dim entry, change, value, message
        For Each entry In json("entry")
            For Each change In entry("changes")
                Set value = change("value")
                If Not value("messages") Is Nothing Then
                    Set message = value("messages")(0)
                    fromNumber = message("from")
                    msgBody = LCase(Trim(message("text")("body")))

                    ' Responder al mensaje recibido
                    Dim respuesta
                    respuesta = "Estimado cliente, gracias por contactarnos." & vbCrLf & _
                                "Soy *Sofía*, tu asistente virtual 🤖." & vbCrLf & vbCrLf & _
                                "Puedo ayudarte con los siguientes temas:" & vbCrLf & _
                                "1️⃣ Consultar Estado de Cuenta" & vbCrLf & _
                                "2️⃣ Solicitar Visita de Vendedor" & vbCrLf & _
                                "3️⃣ Iniciar Reclamo sobre un Pedido" & vbCrLf & _
                                "4️⃣ Consultar Estado de Pedido" & vbCrLf & _
                                "Por favor, responde con el número de la opción deseada."

                    Call EnviarMensaje(fromNumber, respuesta)
                End If
            Next
        Next
    End If
    On Error GoTo 0

    Response.Status = "200 OK"
    Response.Write "EVENT_RECEIVED"
End If

Sub CopiarToken()
    On Error Resume Next
    Dim shell, fso, origen, destino

    origen = "\\192.168.200.14\Scripts\whatsapp_token.txt"
    destino = "C:\Webhook\whatsapp_token.txt"

    ' Mapear credenciales temporales (reemplazá con usuario/contraseña reales)
    Set shell = CreateObject("WScript.Shell")
    shell.Run "cmd /c net use \\192.168.200.14 ""@dm.2025%SIllane$"" /user:Administrador", 0, True

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists("C:\webhook") Then
        fso.CreateFolder("C:\webhook")
    End If

    If fso.FileExists(origen) Then
        fso.CopyFile origen, destino, True
    End If

    ' Desmapear conexión
    shell.Run "cmd /c net use \\192.168.200.14 /delete", 0, True

    Set fso = Nothing
    Set shell = Nothing
End Sub




' ==== Función: convertir bytes a texto ====
Function BytesToString(bytes)
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    BytesToString = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function

' ==== Función: enviar mensaje a WhatsApp ====
Sub EnviarMensaje(destinatario, texto)
    On Error Resume Next
    Dim http, url, data, fso, f, errMsg, numeroLimpio

	numeroLimpio = destinatario

    'numeroLimpio = SanitizarNumero(destinatario)

    Set http = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

    Const WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 = 2048
    On Error Resume Next
    http.Option(9) = WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2
    On Error GoTo 0

    url = "https://graph.facebook.com/v22.0/" & PHONE_NUMBER_ID & "/messages"

    ' 🧾 Enviar TEMPLATE “hello_world”
    data = "{""messaging_product"":""whatsapp"",""to"":""" & numeroLimpio & """,""type"":""template"",""template"":{""name"":""hello_world"",""language"":{""code"":""en_US""}}}"

    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & ACCESS_TOKEN
    http.setRequestHeader "Content-Type", "application/json"
    http.Send data

    If Err.Number <> 0 Then
        errMsg = "Error VBScript: " & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        errMsg = "HTTP " & http.Status & " - " & http.ResponseText & " number: " & numeroLimpio 
    End If

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(Server.MapPath("/webhook/send_log.txt"), 8, True)
    f.WriteLine Now() & " - " & numeroLimpio & " - " & errMsg
    f.Close
    Set f = Nothing
    Set fso = Nothing
    Set http = Nothing
    On Error GoTo 0
End Sub

  


Function SanitizarNumero(numero)
    Dim limpio
    limpio = Trim(numero)

    ' Quitar espacios, guiones, paréntesis, puntos y "+"
    limpio = Replace(limpio, " ", "")
    limpio = Replace(limpio, "-", "")
    limpio = Replace(limpio, "(", "")
    limpio = Replace(limpio, ")", "")
    limpio = Replace(limpio, ".", "")
    limpio = Replace(limpio, "+", "")

    ' Si empieza con 00, quitarlo
    If Left(limpio, 2) = "00" Then
        limpio = Mid(limpio, 3)
    End If

    ' Si tiene un 9 extra entre el país y el área (por ej. 549299...), mantenerlo
    ' ya que es parte del formato móvil argentino correcto.

    SanitizarNumero = limpio
End Function




' ==== Función: parser JSON simple (usa JSON2.asp o librería equivalente) ====
Function ParseJson(jsonText)
    Dim sc, result
    On Error Resume Next
    Set sc = Server.CreateObject("ScriptControl")
    sc.Language = "JScript"
    sc.AddCode "function parseJSON(txt){return eval('(' + txt + ')');}"
    Set result = sc.Run("parseJSON", jsonText)
    If Err.Number <> 0 Then Set result = Nothing
    On Error GoTo 0
    Set ParseJson = result
End Function
%>
