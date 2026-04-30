<%
ExcelFile = Session("FileUploaded")
SQL = "SELECT * FROM [Sheet1$]"
Set ExcelConnection = Server.createobject("ADODB.Connection")
ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelFile & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"";"
SET RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, ExcelConnection 


FOR EACH Column IN RS.Fields
	Response.Write ""
NEXT
IF NOT RS.EOF THEN
	HDRN = RS.Fields(3).value
	'Response.write(HDRN)
	' remuevo la primer linea que tiene el numero de la hoja de ruta, por eso avanzo a la siguiente fila
	RS.movenext

	sSQL = "EXEC dbo.usp_HojaDeRutaPxP_Cargada " & replace(replace(HDRN,",",""),".","")
	Set dbRS=Server.CreateObject("ADODB.Recordset")
	dbRS.open sSQL, dbCon
	If  NOT dbRS.EOF  Then
		if dbRS("Status")=1 THEN
			Response.Write "Ya existe la hoja de ruta"
			msgError = "Ya existe la hoja de ruta"
			RESPONSE.CLEAR
			response.redirect "./errorpxp.asp?errorMessage="&msgError
		End If
	End IF
	
	WHILE NOT RS.eof  
		
		sSQL = "EXEC TMP.usp_HojaDeRutaPxP_Ins " 
		FOR EACH Field IN RS.Fields 
			' tengo que hacer porque trae mas columnas vacias
			if Field.value<>"" then  
				sSQL = sSQL &"'"& Field.value &"',"
			End If 
		NEXT
		'Response.write("FIN<br />")
		sSQL = sSQL &"'"& HDRN &"',"
		ssql = Left(ssql, Len(ssql) - 1)  

		RS.movenext
		'Response.write(sSQL)
		Set dbRS=Server.CreateObject("ADODB.Recordset")
		dbRS.open sSQL, dbCon
		IsParentInactive = False
		IF ERR.NUMBER <> 0 THEN
			RESPONSE.CLEAR
			response.redirect "./error.asp"
		End If		
		
	WEND
	


END IF 
RS.close
ExcelConnection.Close
%>