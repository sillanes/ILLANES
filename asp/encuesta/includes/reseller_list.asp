<%
pre_UserCompanyID = Session("pre_UserCompanyID")
pre_CompanyPermission = Session("CompanyPermission")

sSQL = "EXEC wsp_Resellers_sel"
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	rsCID = dbRS("CompanyID")
	If pre_UserCompanyID=1 Or InStr(pre_CompanyPermission,", "&rsCID&",")>0 Then
		comboList = comboList & "<option value='" & rsCID & "'>" & Left(dbRS("CompanyName"),22) & "</option>"
	End If
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>