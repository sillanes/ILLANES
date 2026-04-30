<%@ Language=VBScript %>
<%'Option Explicit%>
<!--#include file="../includes/gen_func.asp" -->
<%CheckSession "","",""%>
<%
UserCompanyID = Session("UserCompanyID")
Dim dbCon, dbRS, cM, sSQL, comboList
%><!--#include file="../includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="../includes/db_con_open.asp" --><%

If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If

prm = objRequest("prm")

sSQL = "EXEC wsp_preUsers_combo_byCompany  " & prm
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "@@" & Right("00000"&dbRS(0),5) & dbRS(1)
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write Mid(comboList,3)

dbCon.Close
Set dbCon = Nothing
Set objRequest = Nothing
%>