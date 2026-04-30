<%
'ON ERROR RESUME NEXT


If PeriodoID = 0 Then 
	 
	sSQL = "EXEC [pxp].[usp_Periodo_Actual]"
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
sSQL = "EXEC [report].[usp_Controladores_Combo]"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

If EmpleadoID=0 Then
	
	If NOT dbRS.EOF Then
		EmpleadoID = dbRS("EmpleadoID")
	End If 

End If



 
   
%> 


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Controladores: </span></td>
<td>

<form name="FF3" method="post" action="pxpdescargarreclamosxcontrolador.asp">
<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>"> 
<select class="select-input" name="EmpleadoID" class="tboxNoWidth" onchange="chkForm(this.form,1)"> 
<%
  comboList = ""
  sel = ""
  firsttime = 0
  nombrearchivo = ""
  
  While NOT dbRS.EOF
  
		If firsttime=0 THEN
			sSQL = "EXEC  [pxp].[usp_Controladores_Periodo_sel] " & PeriodoID  & "," &  clng("0" & dbRS("EmpleadoID") )
			nombrearchivo = "Periodo_" & PeriodoID  & "_" & dbRS("EmpleadoID")
			firsttime = 1
		End IF
  
		if (dbRS("EmpleadoID") = EmpleadoID) then 
			sel = "Selected"
			sSQL =  "EXEC  [pxp].[usp_Controladores_Periodo_sel] " & PeriodoID  & "," &  clng("0" & dbRS("EmpleadoID") )
			nombrearchivo = "Periodo_" & PeriodoID  & "_" & dbRS("EmpleadoID")
		else 
			sel = ""
		end if 
		
		comboList = comboList & "<option  " & sel & " value='" & dbRS("EmpleadoID") & "'>"  & dbRS("Nombre") & "</option>"
	
	  dbRS.MoveNext
	Wend
	Set dbRS = Nothing
	
 
	
	
	Response.Write comboList
  %>
</select>  
</form>         
</td> 
<td>
<form  action="../app_excel2.asp" method="post" name="FEXCEL"> 
<table border="0" cellpadding="0" cellspacing="0" align="right">
<input type="hidden" name="sSQL" value="<%=sSQL%>">
<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
<tr> 
	<td class="btnExcel" valign="bottom"><input type="image" border="0" src="../images/excel.png" title="Ver en EXCEL"></td>
</tr>
</table>
</form>
</td>
</tr>
</table> 

<% 
Set dbRS = Nothing
%>
