<% 
sSQL = "EXEC pxp.usp_HojaDeRuta_Pendientes_Sel " & hdrid
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%> 
<!-- Botón/ícono de exportación -->
<div style="text-align:right; margin-bottom:10px;">
    <a href=".\includes\exportar_pendientes_excel.asp?hdrid=<%=hdrid%>" target="_blank" title="Exportar a Excel">
        <i class="fa fa-file-excel-o" style="font-size:20px;color:green;"></i> Descargar Excel
    </a>
</div>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%">
	<tbody>
	<tr>
		<td colspan="4" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="4"><span class="formHeader">HOJAS DE RUTA: <%=hdri%></span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="5" width="30"><b>Cliente</b></td> 
		<td height="10" width="30"><b>Razon Social</b></td>
		<td height="5" width="30"><b>Factura</b></td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	<tr class="itemTD11">
		<td><%=dbRS("ClienteID")%></td> 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("RazonSocial")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("FacturaID")%>&nbsp;&nbsp;</td>
 
	</tr>
	
		<%
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="4"> no se encontraron hojas de rutas activas, verifique selección o comuniquese con <a href="mailto:reclamos@illanes.com.ar">reclamos@illanes.com.ar</a>.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>



<% 
Set dbRS = Nothing
%>
