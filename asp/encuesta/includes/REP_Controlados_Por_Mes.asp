<%
'ON ERROR RESUME NEXT
sSQL = "EXEC report.usp_Pxp_Eficiencia_Mes 0,2"
'Response.write(sSQL)
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
		<td colspan="5" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="5"><span class="formHeader">Controlados por mes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Anio</b></td>
		<td height="20" width="30"><b>Periodo</b></td> 
		<td height="20" width="30"><b>Cantidad <br/> Pedidos</b></td>
		<td height="20" width="30"><b>Cantidad <br/> Corregidos</b></td>
		<td height="20" width="30"><b>Cantidad <br/> Reclamos</b></td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("PeriodoAnio")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("PeriodoNombreLargo")%>&nbsp;&nbsp;</td> 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantPedidos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantErroresCorregidos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantReclamos")%>&nbsp;&nbsp;</td>	 
	 </tr>
	
		<% 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="5"> No hay datos</a></td>
		</tr>
		
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
