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
%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="contentTable">
<tr align="center" >
	<td nowrap height="20">&nbsp;<b>CompanyID</b>&nbsp;</td>
	<td nowrap>&nbsp;<b>CompanyName</b>&nbsp;</td>
</tr>
	<%
	sSQL = "EXEC wsp_preUserCompanyPermissions " & prm
	mypagesize = 500
	%>
	<!-- sSQL = <%=sSQL%> -->
	<!--#include file="../includes/paging.asp" -->
	<%If Not_dbRS_EOF Then%>
		<%DO UNTIL dbRS.EOF%>
<tr align="center" bgcolor="" class="">
	<td><%=dbRS("CompanyID2")%></td>
	<td nowrap>&nbsp;&nbsp;<%=dbRS("CompanyName")%>&nbsp;&nbsp;</td>
</tr>
		<%
			howmanyrecs=howmanyrecs+1
			dbRS.MoveNext
		LOOP
		%>
	<%End If%>
	<%Set dbRS = Nothing%>
</table>
<%
dbCon.Close
Set dbCon = Nothing
Set objRequest = Nothing
%>