<%
'ON ERROR RESUME NEXT
sSQL = "EXEC report.usp_Reclamos_rep"
 'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon 
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="8" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="8"><span class="formHeader">RECLAMOS POR DIA</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Fecha</b></td>
		<td height="20" width="30"><b>Abiertos</b></td>
		<td height="20" width="30"><b>Cerrados</b></td>
		<td height="20" width="30"><b>Pendientes</b></td>
		<td height="20" width="30"><b>EnProceso</b></td>
		<td height="20" width="30"><b>TotalReclamos</b></td>
		<td height="20" width="30"><b>TotalFacturas</b></td>
		<td height="20" width="30"><b>%</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="10%"><%=dbRS("FechaFact")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Abiertos")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Cerrados")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Pendientes")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("EnProceso")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalReclamos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalFacturas")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("ReclamosPorcentaje")%>&nbsp;&nbsp;</td>
	</tr>
	
		<% 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="8"> No hay reclamos datos</a>.</td>
		</tr>
		
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
