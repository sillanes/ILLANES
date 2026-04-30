<% 
'Response.write("inicio<br />")
'Response.write(facturaid)
'Response.write(nrocliente)
'Response.write("<br />fin")
sSQL = "EXEC pxp.usp_Buscar_Armador_Controlador " & CLng("0" & facturaid) & "," & CLng("0"&nrocliente)
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
		<td colspan="7 height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="7"><span class="formHeader">FACTURAS - CLIENTES</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="5" width="30"><b>Hoja De Ruta</b></td>
		<td height="5" width="30"><b>Factura</b></td>
		<td height="5" width="30"><b>Cliente</b></td>
		<td height="10" width="30"><b>Razon Social</b></td>
		<td height="5" width="100"><b>Armador</b></td>
		<td height="5" width="30"><b>Controlador</b></td>
		<td height="5" width="40"><b>Accion</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	<tr class="itemTD11">
		<td><%=dbRS("HojaDeRutaNro")%></td>
		<td><%=dbRS("FacturaID")%></td>
		<td><%=dbRS("ClienteNombre")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("RazonSocial")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Armador")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Controlador")%>&nbsp;&nbsp;</td> 
		<td>
			<%If  dbRS("ExisteReclamo")>=1 Then%>
			Ya tiene un reclamo 
			<%else%>
			<a href="javascript:document.FF.doWhat.value='2';chkForm2(document.FF,2,'<%=dbRS("FacturaID")%>','<%=dbRS("ClienteID")%>')" title="Generar Reclamo">
			<img src="../images/upload.png" alt="Generar Reclamo" id="<%=dbRS("FacturaID")%>"></a>
			<%End If%>
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
			<td colspan ="7"> no se encontraron facturas o clientes, verifique selección o comuniquese con <a href="mailto:reclamos@illanes.com.ar">reclamos@illanes.com.ar</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>



<% 
Set dbRS = Nothing
%>
