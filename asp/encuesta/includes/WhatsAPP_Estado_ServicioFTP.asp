<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC python.usp_https_sel " 
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

serviceshttp= dbRS("https") 
 
Set dbRS = Nothing
%>
