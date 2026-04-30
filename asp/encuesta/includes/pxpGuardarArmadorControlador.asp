<% 
'Response.write("inicio<br />")
'Response.write(hdrid)
'Response.write(nrocliente)
'Response.write("<br />fin")
sSQL = "EXEC pxp.usp_Armador_Controlador_upd " & CLng("0" & hdrid) & "," & CLng("0"&nrocliente) & "," & CLng("0"&armadorid) & "," & CLng("0"&controladorid)& "," & CLng("0"&canterrores)
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
