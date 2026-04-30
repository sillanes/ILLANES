<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_Reenvio_Update "& hdrid &","& hojaderutarow&","& agrupacion
'Response.write(sSQL)
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
