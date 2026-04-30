<%
sSQL = "EXEC wsp_Encuesta_Pregunta_sel "
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon 

IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If


If  NOT dbRS.EOF  Then
	aPreguntas = dbRS.GetRows()
End If
Set dbRS = Nothing
%>