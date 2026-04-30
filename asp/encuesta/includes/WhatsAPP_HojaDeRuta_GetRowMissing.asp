<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_GetRow_Missing "& hojaderutarow
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

If  NOT dbRS.EOF  Then
	RDClienteNombre = dbRS("RazonSocial")
	RDTelefono = dbRS("Telefono")
	RDClienteID = dbRS("ClienteID")
End If   
Set dbRS = Nothing
%>
