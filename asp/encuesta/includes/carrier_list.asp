<%
sSQL = "EXEC usp_Carriers_Sel "
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS("CarrierID") & "'>" & Left(dbRS("CarrierName"),24) & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>