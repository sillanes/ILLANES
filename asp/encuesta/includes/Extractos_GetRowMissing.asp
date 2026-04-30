<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Extractos_GetRow_Missing "& extractorow
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

If  NOT dbRS.EOF  Then
	CUIT_Archivo = dbRS("CUIT_Archivo")
End If   
Set dbRS = Nothing
%>
