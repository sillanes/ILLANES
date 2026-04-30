<%
'ON ERROR RESUME NEXT
totalenproceso = 0 
sSQL = "EXEC pxp.usp_Armadores_Controladores_Combo"
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
<td><span class="pagetitle">Armador:</span></td>
<td>
<select id="ArmadorID" class="select-input" name="ArmadorID" class="tboxNoWidth"> 
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

<span class="labelError" id="armadoridmsg"></span>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Errores:</span></td>
<td>
	<input type="text" id="canterrores" name="canterrores" placeholder="9999" minlength="6" required value="<%=canterrores%>">
</td>
</tr>
</table>


<span class="labelError" id="canterrresmsg"></span>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Controlador:</span></td>
<td>
<select id="ControladorID" class="select-input" name="ControladorID" class="tboxNoWidth"> 
<%
  SET dbRS = dbRS.NextRecordset()
  comboList = ""
  While NOT dbRS.EOF
		comboList = comboList & "<option value='" & dbRS("EmpleadoId") & "'>" & dbRS("Nombre") & "</option>"
	
	  dbRS.MoveNext
	Wend
	Set dbRS = Nothing
	Response.Write comboList
  %>
</select>           
</td>
</tr>
</table>

<span class="labelError" id="controladoridmsg"></span>





				
				

<% 
Set dbRS = Nothing
%>
