<%

sSQL = "EXEC usp_Cliente_Sel " & CLng("0"&idCliente)
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False

If NOT dbRS.EOF Then
	ClienteNombre = dbRS("Nombre")   
End If
Set dbRS = Nothing
' Response.write(ClienteNombre)
%>