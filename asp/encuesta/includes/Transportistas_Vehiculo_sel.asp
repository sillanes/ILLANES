<%
'ON ERROR RESUME NEXT
totalenproceso = 0 
sSQL = "EXEC usp_Vehiculos_sel 0"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

If ayudante=0 Then
	TransportistaID2 = 0
End If


%>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Vehiculo:</span></td>
<td>
<select class="select-input" name="VehiculoID" class="tboxNoWidth" onchange="chkForm5(document.FF,3,<%=HojaDeRutaID%>,<%=TransportistaID%>,<%=TransportistaID2%>,<%=ayudante%>)"> 
<%
  comboList = "" 
  sel = ""
  While NOT dbRS.EOF
  
    
	 	if (dbRS("ID") = clng(VehiculoID)) then 
			sel = "Selected" 
		Else 
			sel = ""
		End if
		
		
			comboList = comboList & "<option " & sel & " value='" & dbRS("ID") & "'>" & dbRS("Patente") & "</option>"
	
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
