<%
sSQL = "EXEC ven.[Vendedor_Clientes_Objetivo_Guardar] '" & VendedorID & "'," &  ClienteID & "," &  ObjetivoID & "," & PeriodoID  & ",'"&Session("FileUploaded")&"'"
'Response.write(sSQL)
'Response.end
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>