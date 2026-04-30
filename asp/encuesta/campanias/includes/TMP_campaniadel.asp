<% 
sSQL = "EXEC TMP.usp_Campania_Destinatarios_del"
'Response.write(sSQL)
conn.Execute(sSQL)
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>