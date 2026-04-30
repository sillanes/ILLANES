<%
'ON ERROR RESUME NEXT
sSQL = "EXEC usp_Reclamos_Pendientes_Sel '" & idReclamo & "'"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>
