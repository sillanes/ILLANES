<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"
Response.Buffer = True
Response.Clear

On Error Resume Next

Dim transportistaid, fecha
transportistaid = Trim(Request("transportistaid"))
fecha = Trim(Request("fecha"))

If transportistaid = "" Or fecha = "" Then
    Response.Write "{""error"":""Parametros incompletos""}"
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4   ' Stored Procedure
cmd.CommandText = "tracking.usp_GPS_HojaRuta_Dia"

cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , CLng(transportistaid))
cmd.Parameters.Append cmd.CreateParameter("@Fecha", 200, 1, 10, fecha)

Set rs = cmd.Execute

If Err.Number <> 0 Then
    Response.Write "{""error"":""" & Replace(Err.Description, """", "'") & """}"
    Response.End
End If

If rs.EOF Then
    Response.Write "{""hojaderutaid"":null}"
Else
    Response.Write "{""hojaderutaid"":" & rs(0) & "}"
End If

rs.Close
Set rs = Nothing
Set cmd = Nothing

Response.End
%>
