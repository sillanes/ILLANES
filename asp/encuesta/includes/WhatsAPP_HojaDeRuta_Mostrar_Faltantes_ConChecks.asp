<%
sSQL = "EXEC dbo.[usp_HojaDeRuta_Ver_Faltantes] "& HojaDeRutaID
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If
%>

<!-- Botón para excluir seleccionados -->
<div style="text-align: right; margin-bottom: 10px;">
	<h2>
	<a href="javascript:document.FF.doWhat.value='9';chkFormExcluir(document.FF,9,<%=HojaDeRutaID%>,<%=dbRS("RowID")%>)" title="Excluir Item">Excluir Seleccionados<img src="../images/excluir2.png" alt="Excluir Seleccionados"></a>
	</h2>
</div>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody> 
	<tr>
		<td colspan="11" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr class="tableHeader">
				<td colspan="11" align="left"><span class="formHeader">Hoja De Ruta: <%=HojaDeRutaID%></span></td>
			</tr>
		</table>
		</td>
	</tr>
	</tbody> 

	<tr align="center" class="columnTop">
		<td height="20" width="5%"><input type="checkbox" id="chkAll" onclick="toggleAll(this)"></td>
		<td height="20" width="5%"><b>Cliente</b></td>
		<td height="20" width="15%"><b>Nombre</b></td>
		<td height="20" width="15%"><b>Razon Social</b></td>
		<td height="20" width="5%"><b>Factura</b></td>
		<td height="20" width="10%"><b>Forma de pago</b></td>
		<td height="20" width="5%"><b>Importe Factura</b></td>
		<td height="20" width="5%"><b>Importe a Cobrar</b></td> 
		<td height="20" width="10%"><b>Telefono</b></td> 
		<td height="20" width="10%"><b>Factura</b></td> 
		<td height="20" width="10%"><b>Accion</b></td>  
	</tr>

	<%If NOT dbRS.EOF Then%>
		<%DO UNTIL dbRS.EOF%>
			<tr class="itemTD11">
				<td>
					<input type="checkbox" class="chkRow" value="<%=dbRS("RowID")%>">
				</td>
				<td><%=dbRS("ClienteID")%></td> 
				<td><%=dbRS("ClienteNombre")%></td> 
				<td><%=dbRS("RazonSocial")%></td> 
				<td><%=dbRS("FacturaID")%></td> 
				<td><%=dbRS("FormaPago")%></td> 
				<td>$<%=FormatNumber(dbRS("ImporteFactura"),2)%></td> 
				<td>$<%=FormatNumber(dbRS("ImporteAPagar"),2)%></td> 	 
				<td><%=dbRS("Telefono")%></td> 
				<td><%=dbRS("DirectorioFactura")%></td> 	
				<td>
					<%If dbRS("Telefono") <> "" AND dbRS("DirectorioFactura")<>"" THEN%>
						<a href="javascript:document.FF.doWhat.value='9';chkFormExcluir(document.FF,9,<%=HojaDeRutaID%>,<%=dbRS("RowID")%>)" title="Excluir Item"><img src="../images/excluir2.png" alt="Excluir Item"></a>
						&nbsp;&nbsp;
						<img src="../images/ok.png" alt="Completo">
					<%Else%>
						<a href="javascript:document.FF.doWhat.value='9';chkFormExcluir(document.FF,9,<%=HojaDeRutaID%>,<%=dbRS("RowID")%>)" title="Excluir Item"><img src="../images/excluir2.png" alt="Excluir Item"></a>
						&nbsp;&nbsp;
						<a href="javascript:document.FF.doWhat.value='5';chkForm3(document.FF,5,<%=HojaDeRutaID%>,<%=dbRS("RowID")%>)" title="Modificar Datos"><img src="../images/control.png" alt="Modificar Datos"></a>
						&nbsp;&nbsp;
					<%End IF%>
				</td>
			</tr>
			<%dbRS.MoveNext%>
		<%LOOP%>
	<%Else%>
		<tr class="itemTD11">
			<td colspan="11">No hay hoja de ruta para visualizar</td>
		</tr>
	<%End If%>
</table>

<%Set dbRS = Nothing%>

<!-- JavaScript -->
<script>
function toggleAll(source) {
	const checkboxes = document.querySelectorAll('.chkRow');
	checkboxes.forEach(cb => cb.checked = source.checked);
}

function excluirSeleccionados() {
	const checkboxes = document.querySelectorAll('.chkRow:checked');
	if (checkboxes.length === 0) {
		alert("No hay ítems seleccionados.");
		return;
	}

	if (!confirm("¿Deseás excluir los ítems seleccionados?")) return;

	const rowIDs = Array.from(checkboxes).map(cb => cb.value).join(',');

	document.getElementById("RowIDs").value = rowIDs;
	document.getElementById("formExclusion").submit();
}
</script>