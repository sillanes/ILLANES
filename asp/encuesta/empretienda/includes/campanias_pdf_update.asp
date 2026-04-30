<%
sSQL = "EXEC dbo.usp_Campania_PDF_Upd " & CampaniaID & ",'" & Session("FileUploaded") & "'" 
' Response.write(sSQL)
conn.Execute(sSQL)
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>