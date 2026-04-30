<%
'ON ERROR RESUME NEXT
totalenproceso = 0 
sSQL = "EXEC usp_Transportista_Armadores_Controladores_Combo_V2 '"& idReclamo & "'"
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
<td><span class="pagetitle">Transportista:</span></td>
<td>
<select class="select-input" name="TransportistaID" class="tboxNoWidth"> 
<%
  comboList = ""
  While NOT dbRS.EOF
		comboList = comboList & "<option value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
	
	  dbRS.MoveNext
	Wend 
	Response.Write comboList
  %>
</select>           
</td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Ayudante:</span></td>
<td>
<select class="select-input" name="AyudanteID" class="tboxNoWidth"> 
<%
  SET dbRS = dbRS.NextRecordset()
  comboList = ""
  While NOT dbRS.EOF
		comboList = comboList & "<option value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
	
	  dbRS.MoveNext
	Wend 
	Response.Write comboList
  %>
</select>           
</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Armador:</span></td>
<td>
<select class="select-input" name="ArmadorID" class="tboxNoWidth"> 
<%
  SET dbRS = dbRS.NextRecordset()
  comboList = ""
  While NOT dbRS.EOF
		if dbRS("Selected") = 1 Then 
			comboList = comboList & "<option Selected='True'  value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
		else
			comboList = comboList & "<option  value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
		end if
	
	  dbRS.MoveNext
	Wend 
	Response.Write comboList
  %>
</select>           
</td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Controlador:</span></td>
<td>
<select class="select-input" name="ControladorID" class="tboxNoWidth"> 
<%
  SET dbRS = dbRS.NextRecordset()
  comboList = ""
  While NOT dbRS.EOF
		if dbRS("Selected") = 1 Then 
			comboList = comboList & "<option Selected='True'  value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
		else
			comboList = comboList & "<option  value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
		end if
	  dbRS.MoveNext
	Wend
	Set dbRS = Nothing
	Response.Write comboList
  %>
</select>           
</td>
<td>
<input type="checkbox" id="esReclamoxControl" name="esReclamoxControl" value="1" checked>Error de control<br> 
</td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr class="itemTD11">
	<td><span class="pagetitle">Observaciones:</span></td>
	<td><textarea id="ObservacionesTexto" name="ObservacionesTexto" style="overflow:auto;resize:none" wrap="hard" rows="4" cols="120" maxlength="500"></textarea></td>
</tr>
</table>


<% 
Set dbRS = Nothing
%>
