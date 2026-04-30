<%
'ON ERROR RESUME NEXT
totalenproceso = 0 
sSQL = "EXEC usp_Bancos_Combo"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

IF (BancoID=0) THEN
	BancoID=1
END IF
 

%> 


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm" style="margin-bottom: 15px;">
<tr>
<td><span class="pagetitle">Banco: </span></td>
<td>

<form name="FF1" method="post" action="subirextracto.asp">
<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
<input type="hidden" name="BancoID" value="<%=BancoID%>"> 
<select class="select-input" name="Bancos" class="tboxNoWidth" onchange="chkForm(this.form,0)"> 
<%
  comboList = ""
  comboParentPlanList = ""
  sel = ""
  firsttime = 0
  nombrearchivo = ""
  
  While NOT dbRS.EOF
   
  
		if (dbRS("BancoID") = BancoID AND sel="" ) then 
			sel = "Selected" 			
		else 
			sel = "" 				
		end if 
		
		comboList = comboList & "<option  " & sel & " value='" & dbRS("BancoID") & "'>"  & dbRS("Nombre") & "</option>"
		

 
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

<% 
Set dbRS = Nothing
%>
