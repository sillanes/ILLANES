
<%
If CUIT_Archivo<>"" Then
	cuit=CUIT_Archivo
End If
%>
<div>
	<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
	<tr>
		<td>
		<span>CUIT*</span>
		</td>
		
		<td>
			<input type="text" id="cuitid" name="cuitid" placeholder="99-999999-99" minlength="6" required value="<%=cuit%>"><span class="labelError" id="cuitiderrormsg"></span>
		</td>
	
	</tr>
	<tr>
		<td>
		<span>Cliente*</span>
		</td>
		
		<td>
			<input type="text" id="clientenro" name="clientenro" placeholder="PEPE SRL" minlength="7" required value="<%=rsClienteID%>"><span class="labelError" id="clientenroerrormsg"></span>
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td>
		<span>Vendedor*</span>
		</td>
		
		<td>
			<input type="text" id="vendedornro" name="vendedornro" placeholder="PEPE SRL" minlength="7" required value="<%=rsVendedorID%>"><span class="labelError" id="vendedornroerrormsg"></span>
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td>
		<span>Nro Factura*</span>
		</td>
		
		<td>
			<input type="text" id="facturanro" name="facturanro" placeholder="8792949" minlength="5" required value="<%=rsFacturaNro%>"><span class="labelError" id="facturanroerrormsg"></span>
		</td>
		<td>
		</td>
	</tr>


	</table>
</div>