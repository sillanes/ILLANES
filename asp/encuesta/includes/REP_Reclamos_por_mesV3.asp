<%
'ON ERROR RESUME NEXT
sSQL = "EXEC report.usp_Reclamos_rep_por_mes "
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
		<td colspan="9" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="9"><span class="formHeader">RECLAMOS POR MES</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Periodo</b></td>
		<td height="20" width="30"><b>Abiertos</b></td>
		<td height="20" width="30"><b>Cerrados</b></td>
		<td height="20" width="30"><b>Pendientes</b></td>
		<td height="20" width="30"><b>EnProceso</b></td>
		<td height="20" width="30"><b>TotalReclamos</b></td>
		<td height="20" width="30"><b>TotalFacturas</b></td>
		<td height="20" width="30"><b>%</b></td>
		<td height="20" width="30"><b>Acción</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("FechaFact")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Abiertos")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Cerrados")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Pendientes")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("EnProceso")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalReclamos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalFacturas")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("ReclamosPorcentaje")%>&nbsp;&nbsp;</td>
		<td><a href="./dashboardV3byMonth.asp?PeriodoID=<%=dbRS("PeriodoID")%>" title="Ver estadisticas"><img src="../images/estad.png" alt="Ver estadisticas"></a> </td>
	</tr>
	
		<% 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="9"> No hay datos</a></td>
		</tr>
		
		
	<%
	End If
	
	%>
	
</table>

<% 
 Public Function getNumberOfDays(sdate)

imonth = Right(sdate,2) 
 
Select Case imonth
Case 1, 3, 5, 7, 8, 10, 12
getNumberOfDays = 31
Case 4, 6, 9, 11
getNumberOfDays = 30
Case 2
'logic for checking leap years
If (Year(Date) Mod 4) = 0 Then
If (Year(Date) Mod 100) = 0 Then
If (Year(Date) Mod 400) = 0 Then
getNumberOfDays = 29
Else
getNumberOfDays = 28
End If
Else
getNumberOfDays = 29
End If
Else
getNumberOfDays = 28
End If
End Select
End Function

Set dbRS = Nothing
%>
