<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"

Dim numero
numero = Trim(Request("numero"))

If numero = "" Then
    Response.Write "{""ok"":false,""error"":""Número vacío""}"
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "usp_Transportista_America_Sel"
cmd.Parameters.Append cmd.CreateParameter("@TransportistaCodigoAmerica", 200, 1, 50, numero)

Set rs = cmd.Execute

If rs.EOF Then
    Response.Write "{""ok"":false,""error"":""Transportista no encontrado""}"
Else
    Response.Write "{""ok"":true,""transportistaid"":" & rs("TransportistaID") & _
                   ",""nombre"":""" & Replace(rs("Nombre"),"""","\""") & """}"
End If

rs.Close
Set rs = Nothing
Set cmd = Nothing
Response.End
%>
