<%
sSQL = "EXEC TMP.usp_Reclamos_step_2_upd '"& idReclamo & "','" & Replace(txtresol, "'", "")  & "'"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>
