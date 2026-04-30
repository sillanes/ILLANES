<%

If PeriodoID = 0 Then 
	 
	sSQL = "EXEC [ven].[usp_Periodo_Actual]"
	Set dbRS=Server.CreateObject("ADODB.Recordset")
	dbRS.open sSQL, dbCon
	IsParentInactive = False
	IF ERR.NUMBER <> 0 THEN
		RESPONSE.CLEAR
		response.redirect "./error.asp"
	End If
	
	If NOT dbRS.EOF Then 
		PeriodoID = dbRS("PeriodoID")
	End If 
End If 

 
Set dbRS = Nothing
%>
