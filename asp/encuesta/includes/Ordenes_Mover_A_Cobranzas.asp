<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.[usp_Ordenes_Mover_A_Cobranzas] "& FileID & ",'" & Session("currentUser") & "'"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

hasError= dbRS("hasError")
Mensaje= dbRS("Mensaje")
 
Set dbRS = Nothing
%>
