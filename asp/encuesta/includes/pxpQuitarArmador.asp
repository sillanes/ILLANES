<%
sSQL = "EXEC report.[usp_Nomina_Armadores_upd] '"& armador & "',2" 
'response.write sSQL 
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>