<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True

If CLng("0" & Session("Empleado_UsuarioID")) <= 0 Or CLng("0" & Session("Empleado_EmpleadoID")) <= 0 Then
    Response.Status = "403 Forbidden"
    Response.Write "Acceso denegado."
    Response.End
End If

Dim ruta
ruta = Trim(Request.QueryString("ruta") & "")

If ruta = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "Ruta no informada."
    Response.End
End If

Dim rutaPermitida
rutaPermitida = "\\192.168.200.13\RRHH\Archivos\Vacaciones\"

If LCase(Left(ruta, Len(rutaPermitida))) <> LCase(rutaPermitida) Then
    Response.Status = "403 Forbidden"
    Response.Write "Ruta no permitida."
    Response.End
End If

If LCase(Right(ruta, 4)) <> ".pdf" Then
    Response.Status = "400 Bad Request"
    Response.Write "Archivo inválido."
    Response.End
End If

Dim fso, stm, nombreArchivo
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(ruta) Then
    Response.Status = "404 Not Found"
    Response.Write "El archivo no existe."
    Response.End
End If

nombreArchivo = fso.GetFileName(ruta)

Set stm = Server.CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.LoadFromFile ruta

Response.Clear
Response.ContentType = "application/pdf"
Response.AddHeader "Content-Disposition", "inline; filename=""" & nombreArchivo & """"
Response.BinaryWrite stm.Read

stm.Close
Set stm = Nothing
Set fso = Nothing
Response.End
%>