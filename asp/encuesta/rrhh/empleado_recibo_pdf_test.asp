<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If CLng("0" & Session("Empleado_UsuarioID")) <= 0 Then
    Response.Status = "403 Forbidden"
    Response.Write "Acceso denegado."
    Response.End
End If

Dim rutaArchivo
rutaArchivo = Trim("" & Session("Empleado_TestPDFPath"))

If rutaArchivo = "" Then
    Response.Status = "404 Not Found"
    Response.Write "No hay PDF de prueba generado."
    Response.End
End If

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(rutaArchivo) Then
    Set fso = Nothing
    Response.Status = "404 Not Found"
    Response.Write "El archivo de prueba no existe."
    Response.End
End If

Set fso = Nothing

Dim nombreArchivo
nombreArchivo = Mid(rutaArchivo, InStrRev(rutaArchivo, "\") + 1)

Dim stm
Set stm = Server.CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.LoadFromFile rutaArchivo

Response.Clear
Response.ContentType = "application/pdf"
Response.AddHeader "Content-Disposition", "inline; filename=""" & nombreArchivo & """"
Response.BinaryWrite stm.Read

stm.Close
Set stm = Nothing
Response.End
%>