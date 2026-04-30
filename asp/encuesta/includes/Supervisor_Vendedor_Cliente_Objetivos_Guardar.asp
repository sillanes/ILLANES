<%
sSQL = "EXEC super.[Vendedor_Clientes_Objetivo_Guardar] '" & VendedorID & "'," &  ClienteID & "," &  ObjetivoID & "," & PeriodoID  & ", "  & Acction  & ", '"  & txtdescarte & "'"  
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>