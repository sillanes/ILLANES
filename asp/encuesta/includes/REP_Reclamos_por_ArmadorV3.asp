<%
'ON ERROR RESUME NEXT
sSQL = "EXEC report.usp_Reclamos_Resolucion_Por_Armador_RepV3 " & PeriodoID 
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

dim itotal, icant 

itotal = 0
icant = 0 
pReclamos = 0.00
   
%>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="6" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="6"><span class="formHeader">RECLAMOS POR ARMADOR</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Fecha</b></td>
		<td height="20" width="30"><b>Nombre</b></td>
		<td height="20" width="30"><b>Total reclamos</b></td>
		<td height="20" width="30"><b>Total Armados</b></td>
		<td height="20" width="30"><b>% Reclamos</b></td>  
		<td height="20" width="30"><b>% Armados</b></td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("Fecha")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Nombre")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Total")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Cant_Armados")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("% Reclamos")%>&nbsp;&nbsp;</td>	 	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("% Armados")%>&nbsp;&nbsp;</td>	 
	</tr>
	
		<% 
			itotal = itotal + dbRS("Total")
			icant = icant +  dbRS("Cant_Armados")  
			pReclamos = pReclamos + Cdbl(dbRS("% Reclamos") )
 			dbRS.MoveNext
		LOOP
		%>
		
	<tr class="itemTD11">
		<td colspan="2" align="right">Total</td>
		<td nowrap>&nbsp;&nbsp;<%=itotal%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=icant%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=100%> %&nbsp;&nbsp;</td>	
		<td nowrap><%= ROUND(Cdbl(itotal / icant * 100 ),2) %> %</td>	  
	</tr>	
	
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

Set dbRS = Nothing
%>
