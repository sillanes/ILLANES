<% 
sSQL = "EXEC pxp.usp_HojaDeRuta_Pendientes"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%> 

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%">
	<tbody>
	<tr>
		<td colspan="7" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="7"><span class="formHeader">HOJAS DE RUTA</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="5" width="30"><b>Hoja De Ruta</b></td>
		<td height="5" width="30"><b>Fecha</b></td>
		<td height="5" width="30"><b>Cantidad Clientes</b></td>
		<td height="10" width="30"><b>Cantidad de Facturas</b></td>
		<td height="5" width="30"><b>Facturas Controladas</b></td> 
		<td height="5" width="30"><b>Accion</b></td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	<tr class="itemTD11">
		<td><%=dbRS("HojaDeRutaNro")%></td>
		<td><%=dbRS("Fecha")%></td>
		<td><%=dbRS("Clientes")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Facturas")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Controladas")%>&nbsp;&nbsp;</td>
		<td nowrap>
		<a href="javascript:document.FF.doWhat.value='1';chkVerClientes(document.FF,1,<%=dbRS("HojaDeRutaNro")%>)"  title="Ver Clientes"><img src="../images/eye.png" alt="Ver Clientes" id="<%=dbRS("HojaDeRutaNro")%>"></a>
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
			<td colspan ="7"> no se encontraron hojas de rutas activas, verifique selección o comuniquese con <a href="mailto:reclamos@illanes.com.ar">reclamos@illanes.com.ar</a>.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>



<% 
Set dbRS = Nothing
%>
