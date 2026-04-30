<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC [dbo].[usp_Transportista_sel] 0, "& TransportistaID  
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

If  NOT dbRS.EOF  Then 
	RDNombre =   dbRS("Apellido") &", "& dbRS("Nombre") 
	RDTelefono = dbRS("Telefono") 
End If   
Set dbRS = Nothing
%>
