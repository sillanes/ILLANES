<%@ Language="VBScript" %>
<%
Response.Buffer = True
Response.ContentType = "text/plain"

Dim numero, mensaje
numero = Trim(Request("numero"))
mensaje = Trim(Request("mensaje"))

If numero = "" Or mensaje = "" Then
    Response.Write "ERROR: faltan parámetros."
    Response.End
End If

' Leer token
Dim tokenFile, accessToken, fso, f
tokenFile = "C:\webhook\whatsapp_token.txt"
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(tokenFile) Then
    Set f = fso.OpenTextFile(tokenFile, 1)
    accessToken = Trim(f.ReadAll)
    f.Close
End If
Set f = Nothing : Set fso = Nothing

If accessToken = "" Then
    Response.Write "ERROR: token vacío"
    Response.End
End If

Dim PHONE_NUMBER_ID, url, http, body
PHONE_NUMBER_ID = "556834667519231"
url = "https://graph.facebook.com/v22.0/" & PHONE_NUMBER_ID & "/messages"
body = "{""messaging_product"":""whatsapp"",""to"":""" & numero & """,""type"":""text"",""text"":{""body"":""" & Replace(mensaje, """", "\""") & """}}"

Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.Open "POST", url, False
http.setRequestHeader "Authorization", "Bearer " & accessToken
http.setRequestHeader "Content-Type", "application/json"
http.Send body

Dim s, resp
s = http.Status
resp = http.responseText
Set http = Nothing

If s = 200 Or s = 201 Then
    ' Extraer WAMID
    Dim waid
    waid = ""
    If InStr(resp, """id"":") > 0 Then
        waid = Mid(resp, InStr(resp, """id"":") + 6)
        waid = Left(waid, InStr(waid, """") - 1)
    End If

    ' Registrar en BD
    Dim conn, cmd
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB;Server=192.168.200.13,3306;Database=WhatsAPPApi;UID=whatsapp_user;PWD=Ill@nes.%2025+;"
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = "INSERT INTO sofia.WhatsAPP_API_Historial (WhatsAppID, Numero, Tipo, Estado, Contenido, Origen, Fecha) VALUES (?, ?, ?, ?, ?, ?, GETDATE())"
    cmd.Parameters.Append cmd.CreateParameter("@id", 200, 1, 255, waid)
    cmd.Parameters.Append cmd.CreateParameter("@num", 200, 1, 50, numero)
    cmd.Parameters.Append cmd.CreateParameter("@tipo", 200, 1, 50, "message")
    cmd.Parameters.Append cmd.CreateParameter("@estado", 200, 1, 50, "sent")
    cmd.Parameters.Append cmd.CreateParameter("@cont", 201, 1, 4000, mensaje)
    cmd.Parameters.Append cmd.CreateParameter("@orig", 200, 1, 50, "interno")
    cmd.Execute
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing

    Response.Write "OK"
Else
    Response.Write "ERROR HTTP " & s & ": " & resp
End If
%>
