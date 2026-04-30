<%
sSQL = "EXEC report.[usp_Nomina_Controlador_upd] '"& ControladorID & "',1"

Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>