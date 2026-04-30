<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC dbo.Feriados_sel "
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="2" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="2"><span class="formHeader">Feriados Configurados</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Fecha</td>
		<td height="20" width="30"><b>Acción</b></td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("Fecha")%></td>  
		<td><a href="javascript:document.FF.doWhat.value='2';chkForm2(document.FF,2,'<%=dbRS("Fecha_TO_INT")%>')" title="Borrar"><img src="../images/del.png" alt="Borrar Fecha" id="<%=dbRS("Fecha_TO_INT")%>"></a></td>
	</tr>
	
		<%
			totalPendinetes = totalPendinetes +1
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="2"> No hay feriados configurados</a></td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
