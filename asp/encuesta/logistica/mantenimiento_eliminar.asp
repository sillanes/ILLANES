<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("currentUser") = "" Then Response.End

Dim id
id = Trim(Request("id"))

If id = "" Then Response.End

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "dbo.usp_Vehiculos_Mantenimiento_Eliminar"
cmd.Parameters.Append cmd.CreateParameter("@ID", 3, 1, , CLng(id))

cmd.Execute

Response.Write "OK"

Set cmd = Nothing
conn.Close : Set conn = Nothing
%>
