<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC usp_Reclamos_Pendientes_Sel "
' Response.write(sSQL)
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
		<td colspan="8" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="8"><span class="formHeader">RECLAMOS PENDIENTES</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Fecha<br/>Reclamo </b></td>
		<td height="20" width="30"><b>Número<br/>Reclamo</b></td>
		<td height="20" width="30"><b>Cliente</b></td>
		<td height="20" width="30"><b>Fecha<br/>Factura </b></td>
		<td height="20" width="30"><b>Número<br/>Factura </b></td>
		<td height="20" width="30"><b>Email</b></td>
		<td height="20" width="30"><b>Telefono</b></td>
		<td height="20" width="30"><b>Accion</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("FechaReclamo")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("ReclamoNumero")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Cliente")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("FechaFactura")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("FacturaNumero")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Email")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Telefono")%>&nbsp;&nbsp;</td>
 
		<td><a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,'<%=dbRS("ReclamoNumero")%>')" title="Ver Reclamo"><img src="../images/eye.png" alt="Ver Reclamo" id="<%=dbRS("ReclamoNumero")%>"></a></td>
	</tr>
	
		<%
			totalPendinetes = totalPendinetes +1
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="8"> No hay reclamos pendientes</a>.</td>
		</tr>
		<tr class="itemTD11"><td align="center" colspan ="8"> <input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)"></td></tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
