<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC [dbo].[usp_Reclamos_Mensajes_Save] '"& idReclamo & "'"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

Set dbRS = Nothing
%>
