<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = False%>
<% 
Dim dbCon, dbRS, Conn
Dim dbRS_SORT, ReportName, RepDate, FileName, today, outStr, BlackList, colName, i, fCnt


outStr = ""
BlackList = Request.Form("BlackList")&","
ReportName = Request.Form("ReportName")
If ReportName<>"" Then
	today = date()
	RepDate = year(today) & right("0"&month(today),2) & right("0"&day(today),2)
	FileName = "filename=" & Replace(ReportName," ","") & "_" & RepDate & ".csv"
Else
	FileName = "filename=myfile.csv"
End If


Response.ContentType = "application/csv"
Response.AddHeader "Content-Disposition", FileName

%><!--#include file="db_command_const.asp" --><%
Set dbCon = Server.CreateObject("ADODB.Connection") 
%><!--#include file="db_con_open_administracion.asp" --><%

Server.ScriptTimeout = 300
dbCon.CommandTimeout = 0

Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.cursorlocation=3 'Client=3, Server=2
dbRS.Open Trim(Request.Form("sSQL")), dbCon
fCnt = dbRS.Fields.Count-1
for i = 0 to fCnt
	colName=LCase(dbRS.Fields(i).Name)&","
	if InStr(BlackList,colName)=0 then outStr = outStr & colName
next


outStr = outStr & vbNewLine
while not dbRS.EOF
	for i = 0 to fCnt
		colName=LCase(dbRS.Fields(i).Name)&","
		if InStr(BlackList,colName)=0 then outStr = outStr & Replace(Replace(dbRS.Fields(i).Value&"",",","."), vbNewLine, " ") & ","
	next
	outStr = outStr & vbNewLine
	dbRS.MoveNext
wend

Set dbRS = Nothing
dbCon.Close
Set dbCon = Nothing
Response.Write outStr  
%>
