<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = False%>
<% 
Dim dbCon, dbRS, Conn
Dim dbRS_SORT, ReportName, RepDate, FileName, today, outStr, BlackList, colName, i, fCnt

outStr = ""
 
FileName = "filename=myfile.csv"

Response.ContentType = "application/csv"
Response.AddHeader "Content-Disposition", FileName

%><!--#include file="Includes/db_command_const.asp" --><%
Set dbCon = Server.CreateObject("ADODB.Connection")
Conn = CLng("0"&Request.Form("Conn"))
%><!--#include file="Includes/db_con_open_ventas.asp" --><%
Server.ScriptTimeout = 300
dbCon.CommandTimeout = 0

Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.cursorlocation=3 'Client=3, Server=2
'dbRS.Open Trim(Request.Form("sSQL")), dbCon
dbRS.Open "EXEC super.Supervisor_Vendedores_List 1001,21", dbCon


fCnt = dbRS.Fields.Count-1
for i = 0 to fCnt
	colName=LCase(dbRS.Fields(i).Name)&","
	if InStr(BlackList,colName)=0 then outStr = outStr & colName
next
outStr = outStr & vbNewLine
while not dbRS.EOF
	for i = 0 to fCnt
		colName=LCase(dbRS.Fields(i).Name)&","
		if InStr(BlackList,colName)=0 then outStr = outStr & Replace(Replace(dbRS.Fields(i).Value&"",","," "), vbNewLine, " ") & ","
	next
	outStr = outStr & vbNewLine
	dbRS.MoveNext
wend
 

Set dbRS = Nothing
dbCon.Close
Set dbCon = Nothing
Response.Write outStr
