<%

Dim ExcelConnection, RS, TableRS, ExcelFile, SheetName
ExcelFile = Session("FileUploaded")
'response.write ExcelFile 
'response.write "Session(FileUploaded)->"&Session("FileUploaded") & "*Fin"
'response.write "Session(FileName)->"&Session("FileName") & "*Fin"

Set ExcelConnection = Server.CreateObject("ADODB.Connection")
ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                     "Data Source=" & ExcelFile & ";" & _
                     "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

' Obtener la primera hoja disponible
Set TableRS = ExcelConnection.OpenSchema(20) ' 20 = adSchemaTables
If Not TableRS.EOF Then
    SheetName = TableRS("TABLE_NAME") ' Ej: Sheet1$ o 'Hoja 1$'
Else
    Response.Write "No se encontró ninguna hoja en el Excel."
    Response.End
End If
TableRS.Close
Set TableRS = Nothing

' Hacer SELECT sobre la primera hoja
 
SQL = "SELECT * FROM [" & SheetName & "]"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, ExcelConnection


'
'ExcelFile = Session("FileUploaded")
'response.write ExcelFile
'SQL = "SELECT * FROM [Sheet1$]"
'Set ExcelConnection = Server.createobject("ADODB.Connection")
'ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelFile & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
'SET RS = Server.CreateObject("ADODB.Recordset")
'RS.Open SQL, ExcelConnection 

FOR EACH Column IN RS.Fields
	Response.Write ""
NEXT 
IF NOT RS.EOF THEN
	WHILE NOT RS.eof 
		
		sSQL = "EXEC  [tmp].[usp_Campania_Destinatarios_Ins] " & CampaniaID & ","
		 
		FOR EACH Field IN RS.Fields 
			sSQL = sSQL &"'"& Field.value &"',"
		NEXT
		ssql = Left(ssql, Len(ssql) - 1)  
		RS.movenext
			
 
		conn.Execute sSQL
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