<%
If parentID = "" Then parentID = 0
sSQL = "EXEC wsp_CS_ProblemTicketType_Sel " & parentID
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS(0) & "'>" & dbRS(1) & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>