<%
'ON ERROR RESUME NEXT
totalenproceso = 0 
sSQL = "EXEC report.usp_Nomina_Controladores_Combo"
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
<td><span class="pagetitle">Controlador:</span></td>
<td>
<select class="select-input" name="ControladorID" class="tboxNoWidth"> 
<%
  comboList = ""
  comboParentPlanList = ""
  While NOT dbRS.EOF
		comboList = comboList & "<option value='" & dbRS("EmpleadoID") & "'>" & dbRS("Nombre") & "</option>"
	
	  dbRS.MoveNext
	Wend
	Set dbRS = Nothing
	Response.Write comboList
  %>
</select>           
</td>
<td>
<input type="button" class="btn1" value="Agregar" onClick="chkForm(this.form, 1)">
</td>
</tr>
</table>

<% 
Set dbRS = Nothing
%>
