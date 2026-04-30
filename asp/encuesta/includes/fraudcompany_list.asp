<%
sSQL = "EXEC dbo.wsp_Companies_cmb"
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	rsCID = dbRS("CompanyID")
	comboList = comboList & "<option value='" & rsCID & "'>" & Left(dbRS("CompanyName"),24) & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>