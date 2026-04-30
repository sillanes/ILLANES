<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Ordenes_GetRow_Missing "& FileID & ", "& OrdenID
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
	rsClienteID = dbRS("ClienteID")
	rsVendedorID = dbRS("VendedorID")
	rsFacturaNro = dbRS("FacturaNro")
End If   
Set dbRS = Nothing
%>
