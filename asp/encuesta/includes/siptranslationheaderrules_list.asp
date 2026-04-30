<%
sSQL = "EXEC usp_SipTranslationHeaderRecords_Rules_sel"
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS("ID") & "'>" & dbRS("HeaderRule") & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>