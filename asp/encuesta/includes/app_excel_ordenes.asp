<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<%
Dim dbCon, dbRS, ReportName, RepDate, FileName, today
Dim BlackList, colName, i, fCnt

BlackList = Request.Form("BlackList") & ","
ReportName = Request.Form("ReportName")

If ReportName <> "" Then
    today = Date()
    RepDate = Year(today) & Right("0" & Month(today),2) & Right("0" & Day(today),2)
    FileName = Replace(ReportName," ","") & "_" & RepDate & ".xls"
Else
    FileName = "myfile.xls"
End If

' Cabeceras para descarga en Excel
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & FileName
Response.CharSet = "UTF-8"
Response.CodePage = 65001

' --- Conexión ---
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="db_con_open_administracion.asp" --><%
Server.ScriptTimeout = 300
dbCon.CommandTimeout = 0

Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.CursorLocation = 3
dbRS.Open Trim(Request.Form("sSQL")), dbCon

fCnt = dbRS.Fields.Count - 1

' --- Tabla HTML ---
Response.Write "<table border='1'>"
Response.Write "<tr>"
For i = 0 To fCnt
    colName = LCase(dbRS.Fields(i).Name)
    If InStr(BlackList, colName & ",") = 0 Then Response.Write "<th>" & colName & "</th>"
Next
Response.Write "</tr>"

While Not dbRS.EOF
    Response.Write "<tr>"
    For i = 0 To fCnt
        colName = LCase(dbRS.Fields(i).Name)
        If InStr(BlackList, colName & ",") = 0 Then
            Response.Write "<td>" & Replace(dbRS.Fields(i).Value & "", vbNewLine, " ") & "</td>"
        End If
    Next
    Response.Write "</tr>"
    dbRS.MoveNext
Wend

Response.Write "</table>"

dbRS.Close
Set dbRS = Nothing
dbCon.Close
Set dbCon = Nothing
%>
