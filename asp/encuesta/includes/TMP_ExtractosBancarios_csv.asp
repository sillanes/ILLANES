<%
ExcelFile = Session("FileUploaded")
 
sSQL = "EXEC TMP.usp_ExtractosBancarios_Ins " & BancoID & ",'"&Session("FileUploaded")&"'"
'response.write sSQL 
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If		

 
%>