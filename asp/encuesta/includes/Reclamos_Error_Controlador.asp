<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Reclamos_Error_Controlador_Combo"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

showcombo = ""

IF fromwhat=2 Then
	showcombo = "disabled"
End IF

   
%> 


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">¿Es un reclamo debido a falta de control?</span></td>
<td> 
 
<select class="select-input" name="ReclamoxControl" class="tboxNoWidth" <%=showcombo%> onchange="chkForm(this.form,1)"> 
<%
  comboList = "" 
  sel = "" 
  
  While NOT dbRS.EOF
  
	 	if (dbRS("Id") = ReclamoxControl) then 
			sel = "Selected"
		Else 
			sel = ""
		End if
		comboList = comboList & "<option  " & sel & " value='" & dbRS("ID") & "'>"  & dbRS("Nombre") & "</option>"
	
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
