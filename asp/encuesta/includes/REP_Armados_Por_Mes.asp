<%
'ON ERROR RESUME NEXT
sSQL = "EXEC report.usp_Pxp_Eficiencia_Mes 0,1"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

IF Session("currentUser") = "admin" or Session("currentUser") = "rrhh" Then 
   
%>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="8" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="8"><span class="formHeader">Armadores por mes</span></td>
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
		<td height="20" width="30"><b>Cantidad <br/> Errores</b></td> 
		<td height="20" width="30"><b>Eficiencia %</b></td>
		<td height="5" width="30"><b>Premio</b></td> 
		<td height="5" width="30"><b>Total</b></td> 			
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("PeriodoAnio")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("PeriodoNombre")%>&nbsp;&nbsp;</td> 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantPedidos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantErrores")%>&nbsp;&nbsp;</td>	 
		<td nowrap>&nbsp;&nbsp;<%=FormatNumber(dbRS("Eficiencia"),2)%>&nbsp;&nbsp;%</td>	
		<td nowrap>&nbsp;&nbsp;$ <%=dbRS("Premio")%>&nbsp;&nbsp;</td>	
		<td align="right" nowrap>&nbsp;&nbsp;$ <%=FormatNumber(dbRS("Total"),2)%>&nbsp;&nbsp;</td>	  
	 </tr>
	
		<% 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="8"> No hay datos</a></td>
		</tr>
		
		
	<%
	End If
	
	%>
	
</table>

<% 
Else
%>
<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="6" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="6"><span class="formHeader">Armadores por mes</span></td>
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
		<td height="20" width="30"><b>Cantidad <br/> Errores</b></td> 
		<td height="20" width="30"><b>Eficiencia %</b></td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("PeriodoAnio")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("PeriodoNombre")%>&nbsp;&nbsp;</td> 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantPedidos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CantErrores")%>&nbsp;&nbsp;</td>	 
		<td nowrap>&nbsp;&nbsp;<%=FormatNumber(dbRS("Eficiencia"),2)%>&nbsp;&nbsp;%</td>	  
	 </tr>
	
		<% 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="6"> No hay datos</a></td>
		</tr>
		
		
	<%
	End If
	
	%>
	
</table>
<%
End If
%>


<% 
Set dbRS = Nothing
%>
