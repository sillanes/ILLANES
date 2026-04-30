<%
sSQL = "EXEC wsp_Regions_sel " & CLng("0" & Country)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS("RegionID") & "'>" & dbRS("Description") & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>