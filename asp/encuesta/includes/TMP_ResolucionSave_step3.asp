<%
sSQL = "EXEC TMP.usp_Reclamos_step_3_upd '"& idReclamo & "','"&Session("FileUploaded")&"'"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>