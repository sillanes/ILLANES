<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.ContentType = "application/json"
Response.CharSet = "UTF-8"

Dim transportistaid
transportistaid = Request("transportistaid")

If transportistaid = "" Then
    Response.Write "{""error"":""TransportistaID requerido""}"
    Response.End
End If

Dim fechaHoy
fechaHoy = Year(Date()) & "-" & Right("0"&Month(Date()),2) & "-" & Right("0"&Day(Date()),2)

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "tracking.usp_Transportista_HojaDeRuta_Dia"

cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , CLng(transportistaid)) 
Set rs = cmd.Execute

Response.Write "{""fecha"":""" & fechaHoy & """,""hojas"":["

Dim first : first = True
Do While Not rs.EOF
    If Not first Then Response.Write ","
    first = False

    Response.Write "{"
    Response.Write """id"":" & rs("HojaDeRutaID")  
    Response.Write "}"

    rs.MoveNext
Loop

Response.Write "]}"

rs.Close
Set rs = Nothing
Set cmd = Nothing
%>
