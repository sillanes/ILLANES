<%
'ON ERROR RESUME NEXT
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

'response.write "PeriodoID>"&PeriodoID

totalenproceso = 0 
sSQL = "EXEC ven.Periodos_Combo"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%> 


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Periodo:</span></td>
	<td>  
	 
	<select class="select-input" name="PeriodoID" class="tboxNoWidth" onchange="chkForm(this.form,0)"> 
	<%
	  comboList = ""
	  sel = ""
	  firsttime = 0
	  nombrearchivo = ""
	  
	  While NOT dbRS.EOF
	  
			If firsttime=0 THEN
 
				firsttime = 1
			End IF
	  
			if (dbRS("RowID") = PeriodoID) then 
				sel = "Selected"
 
			else 
				sel = ""
			end if 
			
			comboList = comboList & "<option  " & sel & " value='" & dbRS("RowID") & "'>"  & dbRS("Descripcion") & "</option>"
			
		  dbRS.MoveNext
		Wend
		Set dbRS = Nothing
		
		
		Response.Write comboList
	  %>
	</select> 
	  
	</td> 
</tr>
</table>
<% 
Set dbRS = Nothing
%>
