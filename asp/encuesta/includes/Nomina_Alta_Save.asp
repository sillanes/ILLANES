<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC [dbo].[usp_Nomina_ins] '"& apellido & "','" & nombre & "'," & esLogistica & "," & esTransportista 
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
