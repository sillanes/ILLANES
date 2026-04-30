<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = False%>
<%
If Session("pre_userid")&"" <> "" Then
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

	%><!--#include file="Includes/db_command_const.asp" --><%
	Set dbCon = Server.CreateObject("ADODB.Connection")
	Conn = CLng("0"&Request.Form("Conn"))
	If Conn=1 Then
		%><!--#include file="Includes/db_con_open_prepaid.asp" --><%
	ElseIf Conn=3 Then
		%><!--#include file="Includes/db_con_open.asp" --><%
	ElseIf Conn=4 Then
		%><!--#include file="Includes/db_con_open_post.asp" --><%
	Else
		%><!--#include file="Includes/db_con_open_report.asp" --><%
	End If
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
			if InStr(BlackList,colName)=0 then outStr = outStr & Replace(Replace(dbRS.Fields(i).Value&"",","," "), vbNewLine, " ") & ","
		next
		outStr = outStr & vbNewLine
		dbRS.MoveNext
	wend

'	Dim xmlDoc, attrNode, rowNode, rows, colStr
'	Set xmlDoc = Server.CreateObject("Msxml2.DOMDocument.6.0")
'	Set dbRS=Server.CreateObject("ADODB.Recordset")
'	dbRS.open Trim(Request.Form("sSQL")), dbCon
'	dbRS.Save xmlDoc, 1

'	If xmlDoc.hasChildNodes() Then
'		xmlDoc.setProperty "SelectionNamespaces", "xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'"
'		xmlDoc.setProperty "SelectionLanguage", "XPath"
'		fCnt = 1
'		set rows = xmlDoc.selectNodes("//rs:data/z:row")
'		for each rowNode in rows
'			colStr = ""  
'			outStr = ""
'			for each attrNode in rowNode.selectNodes("@*")
'				if InStr(BlackList,attrNode.nodeName)=0 then 
'					colStr = colStr & attrNode.nodeName & ","
'					outStr = outStr & Replace(Replace(attrNode.nodeTypedValue&"",","," "), vbNewLine, " ") & ","
'				end if
'			next
'			If fCnt = 1 Then
'				fCnt = 0
'				Response.Write colStr & vbCRLF  
'			End If
'			Response.Write outStr & vbCRLF
'		next
'	End If
'	Set xmlDoc = Nothing

	Set dbRS = Nothing
	dbCon.Close
	Set dbCon = Nothing
	Response.Write outStr
End If%>