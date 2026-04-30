<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("currentUser") = "" Then Response.End

Dim vid, comentario, usuario
vid        = Trim(Request.Form("vehiculoID"))
comentario = Trim(Request.Form("comentario"))
usuario    = Session("currentUser")

If vid = "" Or comentario = "" Then
    Response.End
End If

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "dbo.usp_Vehiculos_Mantenimiento_Guardar"

cmd.Parameters.Append cmd.CreateParameter("@VehiculoID", 3, 1, , CLng(vid))
cmd.Parameters.Append cmd.CreateParameter("@Usuario", 200, 1, 50, usuario)
cmd.Parameters.Append cmd.CreateParameter("@Comentario", 203, 1, 500, comentario)

cmd.Execute

Response.Write "OK"

Set cmd = Nothing
conn.Close : Set conn = Nothing
%>
