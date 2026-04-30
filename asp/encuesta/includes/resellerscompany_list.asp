<%
sSQL = "EXEC wsp_ResellersCompanies_sel " & CLng("0"&Reseller)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS("CompanyID") & "'> - " & Left(dbRS("CompanyName"),24) & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>