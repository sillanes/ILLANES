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
	<form name="FF" method="post" action="vendescargarestadoporperiodos.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">  
	 
	<select class="select-input" name="PeriodoID" class="tboxNoWidth" onchange="chkForm(this.form,1)"> 
	<%
	  comboList = ""
	  sel = ""
	  firsttime = 0
	  nombrearchivo = ""
	  
	  While NOT dbRS.EOF
	  
			If firsttime=0 THEN
				sSQL = "EXEC super.Supervisor_Vendedores_List " &"'"& SupervisorID & "'," & PeriodoID 
				nombrearchivo = "Estado_General_" & PeriodoID  & "_" & SupervisorID
				firsttime = 1
			End IF
	  
			if (dbRS("RowID") = PeriodoID) then 
				sel = "Selected"
				sSQL =  "EXEC super.Supervisor_Vendedores_List " &"'"& SupervisorID & "'," & PeriodoID 
				nombrearchivo = "Estado_General_" & PeriodoID  & "_" & SupervisorID
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
	 
	</form>       
	</td>
	<td>

	<form action="app_excel3.asp" method="post" name="FEXCEL">  
	<input type="hidden" name="sSQL" value="EXEC [ven].[usp_Periodo_Actual]">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
 
		<span class="btnExcel" valign="bottom"><input type="image" border="0" src="../images/excel.png" title="Ver en EXCEL"></span>
   
	</form>

	</td>
</tr>
</table>
<% 
Set dbRS = Nothing
%>
