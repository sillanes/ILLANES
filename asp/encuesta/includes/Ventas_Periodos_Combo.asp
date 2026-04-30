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

'ON ERROR RESUME NEXT 
sSQL = "EXEC ven.Periodos_Combo"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If




   
%> 

<div style="overflow-x: auto;">
<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Periodos: </span></td>
<td>
<form name="FF" method="post" action="vendedorcarga.asp">
<input type="hidden" name="doWhat" value="<%=doWhat%>">  
 
<select class="select-input" name="PeriodoID" class="tboxNoWidth" onchange="chkForm(this.form,1)"> 
<%
  comboList = ""
  comboParentPlanList = ""
  sel = "" 
  
  While NOT dbRS.EOF
  
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
</form>       
</td>
</tr>
</table>
</div>
<% 
Set dbRS = Nothing
%>
