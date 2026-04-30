<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.ContentType = "application/json"
Response.Buffer = True
Response.Clear

On Error Resume Next

' =========================
' LECTURA PARÁMETROS
' =========================
Dim transportistaid, androidid
transportistaid = Trim(Request("transportistaid"))
androidid       = Trim(Request("androidid"))


If transportistaid = "" Or androidid = "" Then
    Response.Write "{""estado"":""ERROR"",""mensaje"":""Faltan parámetros""}"
    Response.End
End If

' =========================
' EJECUTAR SP
' =========================
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4 ' Stored Procedure
cmd.CommandText = "tracking.usp_Transportista_ValidarDispositivo"

cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , CLng(transportistaid))
cmd.Parameters.Append cmd.CreateParameter("@AndroidID", 200, 1, 255, androidid)

Set rs = cmd.Execute

If Err.Number <> 0 Then
    Response.Write "{""estado"":""ERROR"",""mensaje"":""" & Replace(Err.Description, """", "'") & """}"
    Response.End
End If

' =========================
' RESPUESTA
' =========================
If rs.EOF Then
    Response.Write "{""estado"":""BLOQUEADO""}"
Else
    Dim estado
    estado = UCase(Trim(rs("Estado")))
    If estado <> "OK" And estado <> "BLOQUEADO" Then estado = "BLOQUEADO"

    Response.Write "{""estado"":""" & estado & """}"
End If

rs.Close
Set rs = Nothing
Set cmd = Nothing

Response.End
%>
