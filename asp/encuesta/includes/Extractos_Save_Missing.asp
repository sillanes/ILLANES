<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Extractos_Save_Missing "& Extracto &","& extractorow&",'"& cuit&"','"& razonsocial &"'" & ", " & chksincuit 
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
