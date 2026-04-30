<%
sSQL = "EXEC [dbo].[usp_Reclamos_Clasificacion_upd] '"& idReclamo & "',"   & ResolucionMotivoID   & " ,"  & ReclamoxControl  
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>