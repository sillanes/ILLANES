<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "tracking.usp_GPS_Recorrido_Sel"

' ===============================
' PARAMETROS COMO VARCHAR
' ===============================
cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , CInt(Request("transportistaid")))
cmd.Parameters.Append cmd.CreateParameter("@Fecha", 200, 1, 10, Request("fecha"))

Set rs = cmd.Execute

Response.Write "["

Do While Not rs.EOF
    Response.Write "{""lat"":" & Replace(rs("Latitud"),",",".") & _
                   ",""lng"":" & Replace(rs("Longitud"),",",".") & _
                   ",""fecha"":""" & rs("FechaHoraGPS") & """}"
    rs.MoveNext
    If Not rs.EOF Then Response.Write ","
Loop

Response.Write "]"
%>
