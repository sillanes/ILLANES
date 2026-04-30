<% 
'Response.write("inicio<br />")
'Response.write(hdrid)
'Response.write(nrocliente)
'Response.write("<br />fin")
sSQL = "EXEC pxp.usp_HDR_Close " & CLng("0" & hdrid)  
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>  


<% 
Set dbRS = Nothing
%>
