<%
'ON ERROR RESUME NEXT
sSQL = "EXEC dbo.usp_Buscar_Reclamo " & idCliente & "," & idFactura
'Response.write("EXEC dbo.usp_Buscar_Reclamo ")
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>


	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody>
		<tr>
			<td colspan="10" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="10"><span class="formHeader">Datos</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" width="20"><b>Cliente</b></td>
			<td height="20" width="10"><b>Fecha Factura</b></td>
			<td height="20" width="20"><b>Fecha Reclamo</b></td>
			<td height="20" width="10"><b>Nro Factura</b></td>
			<td height="20" width="10"><b>Email</b></td>
			<td height="20" width="10"><b>Telefono</b></td>
			<td height="20" width="10"><b>Estado</b></td>
			<td height="20" width="10"><b>Motivo</b></td>
			<td height="20" width="20"><b>Reclamo</b></td>
			<td height="20" width="5"><b>Acción</b></td>
		</tr>
 
 	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			

		<tr  align="left" class="itemTD11">
			<td align="left" width="20%"><%=dbRS("Cliente")%></td>
			<td align="left" ><%=dbRS("FechaFactura")%></td>
			<td align="left" width="10%" ><%=dbRS("FechaReclamo")%></td>
			<td><%=dbRS("FacturaNumero")%></td>
			<td align="left" width="10%"><%=trim(dbRS("email"))%></td>
			<td><%=dbRS("telefono")%></td>
			<td><%=dbRS("MotivoDescripcion")%></td>
			<td><%=dbRS("statusdes")%></td>
			<td align="left" width="20%" ><%=dbRS("ReclamoNumero")%></td>
			<td><a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,2,'<%=dbRS("ReclamoNumero")%>')" title="Ver Reclamo"><img src="../images/eye.png" alt="Ver Reclamo" id="<%=dbRS("ReclamoNumero")%>"></a></td>
			
		</tr>
				 
		<%
			dbRS.MoveNext
		LOOP
		%>				 
				
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="8"> No se encontraron datos</td>
		</tr>
		
	<%
	End If
	
	%>
	</table>
	
	
  