<%
ExcelFile = Session("FileUploaded")
SQL = "SELECT * FROM [Sheet1$]"
Set ExcelConnection = Server.createobject("ADODB.Connection")
ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelFile & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
SET RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, ExcelConnection 

FOR EACH Column IN RS.Fields
	Response.Write ""
NEXT 
IF NOT RS.EOF THEN
	WHILE NOT RS.eof 
		
		sSQL = "EXEC TMP.usp_HojaDeRuta_Ins "
		 
		FOR EACH Field IN RS.Fields 
			sSQL = sSQL &"'"& Field.value &"',"
		NEXT
		ssql = Left(ssql, Len(ssql) - 1)  
		RS.movenext
		'response.write sSQL
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