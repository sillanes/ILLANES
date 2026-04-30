<% 
sSQL = "EXEC dbo.usp_Transportista_DesVincular_HDR "& HojaDeRutaID 
'Response.write(sSQL)
'response.end
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

hasError= dbRS("hasError")
ErrorMessage= dbRS("ErrorMessage")
 
Set dbRS = Nothing
%>