<%
'ON ERROR RESUME NEXT
totalenproceso = 0 
sSQL = "EXEC pxp.Periodos_Combo"
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
<td><span class="pagetitle">Periodos: </span></td>
<td>

<form name="FF" method="post" action="pxpdescargarperiodo.asp">
<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>"> 
<select class="select-input" name="Periodos" class="tboxNoWidth" onchange="chkForm(this.form,0)"> 
<%
  comboList = ""
  comboParentPlanList = ""
  sel = ""
  firsttime = 0
  nombrearchivo = ""
  
  While NOT dbRS.EOF
  
		If firsttime=0 THEN
			sSQL = "EXEC [pxp].[usp_Armadores_Controladores_Periodo_sel] " &  clng("0" & dbRS("RowID") )
			nombrearchivo = "Periodo_" & dbRS("Descripcion")
			firsttime = 1
		End IF
  
		if (dbRS("RowID") = PeriodoID) then 
			sel = "Selected"
			sSQL = "EXEC [pxp].[usp_Armadores_Controladores_Periodo_sel] " &  clng("0" & dbRS("RowID") )
			nombrearchivo = "Periodo_" & dbRS("Descripcion")
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
<table border="0" cellpadding="0" cellspacing="0" align="right">
<form action="../app_excel2.asp" method="post" name="FEXCEL">
<input type="hidden" name="sSQL" value="<%=sSQL%>">
<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
<tr> 
	<td class="btnExcel" valign="bottom"><input type="image" border="0" src="../images/excel.png" title="Ver en EXCEL"></td>
</tr></form></table>
</td>
</tr>
</table> 

<% 
Set dbRS = Nothing
%>
