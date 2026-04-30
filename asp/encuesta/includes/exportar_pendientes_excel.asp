<%@ Language="VBScript" %> 
<%
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 

Dim hdrid
hdrid = Request.QueryString("hdrid")

sSQL = "EXEC pxp.usp_HojaDeRuta_Pendientes_Sel " & hdrid
' response.write sSQL
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.Open sSQL, dbCon

' Configuración para que el navegador lo descargue como Excel
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=HojasDeRuta_" & hdrid & ".xls"
Response.Write "<table border='1'>"
Response.Write "<tr>"
Response.Write "<th>Cliente</th>" 
Response.Write "<th>Razon Social</th>"
Response.Write "<th>Factura</th>" 
Response.Write "</tr>"

If Not dbRS.EOF Then
    Do Until dbRS.EOF
        Response.Write "<tr>"
        Response.Write "<td>" & dbRS("ClienteID") & "</td>" 
        Response.Write "<td>" & dbRS("RazonSocial") & "</td>"
        Response.Write "<td>" & dbRS("FacturaID") & "</td>"
        Response.Write "</tr>"
        dbRS.MoveNext
    Loop
Else
    Response.Write "<tr><td colspan='4'>Sin datos disponibles</td></tr>"
End If

Response.Write "</table>"

dbRS.Close
Set dbRS = Nothing
dbCon.Close
Set dbCon = Nothing
%>
