<% 
sSQL = "EXEC report.usp_Nomina_Armadores_sel "
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%> 

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 50%">
	<tbody>
	<tr>
		<td colspan="2" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="2"><span class="formHeader">Armadores</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="5" width="5"><b>Nommbre</b></td> 
		<td height="5" width="5"><b>Accion</b></td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	<tr class="itemTD11">
		<td><%=dbRS("Nombre")%></td>  
		<td nowrap>
				<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,2,<%=dbRS("EmpleadoID")%>)"  title="Quitar armador"><img src="../images/del.png" alt="Quitar armador" id="<%=dbRS("EmpleadoID")%>"></a>
		</td>
	</tr>
	
		<%
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="2"> no se encontraron armadores.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>



<% 
Set dbRS = Nothing
%>
