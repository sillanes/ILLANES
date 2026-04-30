<%

' @ReclamoID VARCHAR(25),
' @email VARCHAR(100),
' @archivoadjunto VARCHAR(255) = NULL
 
sSQL = "EXEC usp_Reclamos_Pendientes_update '"& idReclamo & "','" & Replace(txtresol, "'", "")  & "'" & "," & "'" & Session("FileUploaded") & "'"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>
