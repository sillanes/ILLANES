<%
ExcelFile = Session("FileUploaded")
SQL = "SELECT * FROM DetalleMovimientosHSBC.csv"

response.write ExcelFile		
Set ExcelConnection = Server.createobject("ADODB.Connection")
ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\Reclamos\Extractos';Extended Properties=""text;HDR=yes;FMT=Delimited(,);MaxScanRows=0;IMEX=0"";"
SET RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, ExcelConnection 

FOR EACH Column IN RS.Fields
	Response.Write ""
NEXT 
IF NOT RS.EOF THEN
	WHILE NOT RS.eof 
		
		sSQL = "EXEC TMP.usp_ExtractosBancarios_Ins "
		response.write RS.Fields(5)
		response.write "<br/>Start"
		FOR EACH Field IN RS.Fields 
		
			response.write Field.value
		
			sSQL = sSQL &"'"& Field.value &"',"
		NEXT
		ssql = Left(ssql, Len(ssql) - 1)  
		RS.movenext
		
		response.write "<br/>end"
		
		response.write ssql
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