<%
'ON ERROR RESUME NEXT
sSQL = "EXEC [report].[usp_Reclamos_Resolucion_Por_Motivos_Rep] "
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
		<td colspan="11" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="11"><span class="formHeader">RECLAMOS POR MOTIVOS</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 
	<tr align="center" class="columnTop">
		<td height="20" width="5"><b>Fecha</b></td>
		<td height="20" width="10"><b>FALTANTE DE MERCADERIA SIN STOCK</b></td>
		<td height="20" width="10"><b>FALTANTE DE MERCADERIA ENVIADA</b></td>
		<td height="20" width="10"><b>FALTANTE DE MERCADERIA</b></td>
		<td height="20" width="5"><b>ERROR UNIDAD DE MEDIDA</b></td>
		<td height="20" width="5"><b>MERCADERIA EQUIVOCADA</b></td>
		<td height="20" width="5"><b>MERCADERIA DAÑADA</b></td>
		<td height="20" width="5"><b>DEVOLUCION DE MERCADERIA</b></td>
		<td height="20" width="5"><b>SIN DEFINIR</b></td>
		<td height="20" width="5"><b>OTRO</b></td> 
		<td height="20" width="5"><b>TOTALES</b></td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("Fecha")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("FALTANTE MERCADERIA SIN STOCK")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("FALTANTE MERCADERIA ENVIADA")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("FALTANTE MERCADERIA")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("ERROR UNIDAD DE MEDIDA")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("MERCADERIA EQUIVOCADA")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("MERCADERIA DANADA")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("DEVOLUCION DE MERCADERIA")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("SIN DEFINIR")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("OTRO")%>&nbsp;&nbsp;</td> 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TOTAL")%>&nbsp;&nbsp;</td> 
	</tr>
	
		<% 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="11"> No hay datos</a></td>
		</tr>
		
		
	<%
	End If
	
	%>
	
</table>

<%  

Set dbRS = Nothing
%>
