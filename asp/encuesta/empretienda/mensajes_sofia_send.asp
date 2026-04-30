<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Server.ScriptTimeout = 120

' ==========================================================
' mensajes_sofia_send.asp
' Envía un mensaje a través del webhook Flask (puerto 444)
' ==========================================================

Dim numero, mensaje, url, jsonBody, http, respuesta

numero = Trim(Request("numero"))
mensaje = Trim(Request("mensaje"))

If numero = "" Or mensaje = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "ERROR: faltan parámetros (numero o mensaje)"
    Response.End
End If

' ============================
' Construir JSON a enviar
' ============================
jsonBody = "{""to"":""" & numero & """,""message"":""" & Replace(mensaje, """", "\""") & """}"

' ============================
' Endpoint Flask (Webhook activo)
' ============================
url = "https://illanes-encuesta.ddns.net:444/send_message"

On Error Resume Next
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Timeout de conexión y envío
http.setTimeouts 10000, 10000, 10000, 10000

' Abrir conexión segura con SSL
http.Open "POST", url, False

' Cabeceras necesarias
http.setRequestHeader "Content-Type", "application/json"

' Enviar mensaje
http.Send jsonBody

If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "❌ Error enviando mensaje: " & Err.Description
    On Error GoTo 0
    Response.End
End If
On Error GoTo 0

' ============================
' Mostrar respuesta
' ============================
Dim respText, respStatus
respStatus = http.Status
respText = http.responseText

If respStatus >= 200 And respStatus < 300 Then
    ' OK → mensaje enviado correctamente
    Response.Write "✅ Mensaje enviado correctamente (" & respStatus & "): " & respText
Else
    Response.Write "⚠ Error en la respuesta del servidor (" & respStatus & "): " & respText
End If

' ============================
' Log local (para depuración)
' ============================
Dim fso, logFile, logPath
logPath = Server.MapPath("/webhook/logs/mensajes_sofia_log.txt")

Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Server.MapPath("/webhook/logs/")) Then
    fso.CreateFolder(Server.MapPath("/webhook/logs/"))
End If

Set logFile = fso.OpenTextFile(logPath, 8, True, 0)
logFile.WriteLine Now() & " | " & numero & " | " & Replace(mensaje, vbCrLf, " ") & " | RESP: " & respStatus & " " & respText
logFile.Close

Set logFile = Nothing
Set fso = Nothing
Set http = Nothing
%>
