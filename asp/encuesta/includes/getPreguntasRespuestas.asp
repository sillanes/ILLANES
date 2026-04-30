<%
sSQL = "EXEC wsp_Encuesta_Pregunta_Respuesta_sel 1"
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon 

If  NOT dbRS.EOF  Then
	aRespuestas = dbRS.GetRows()
End If 
Set dbRS = Nothing
%>