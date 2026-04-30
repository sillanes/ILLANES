<%
sSQL = "EXEC wsp_CompanyAgents_combo " & CLng("0"&Company)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS("AgentID") & "'>" & dbRS("Name") & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>