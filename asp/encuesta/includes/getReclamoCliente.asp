<%
'ON ERROR RESUME NEXT
sSQL = "EXEC dbo.usp_Reclamos_Sel '" & idReclamo & "'"  
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>

<h1>Reclamo: <%=idReclamo%></h1>
<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
	<tr>
		<td colspan="4" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="4"><span class="formHeader">Datos seleccionados</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="left" class="columnTop">
		<td height="20" width="60"><b>Cliente</b></td>
		<td height="20" width="30"><b>Fecha</b></td>
		<td height="20" width="30"><b>Nro Factura</b></td>
		<td height="20" width="30"><b>Motivo Reclamo</b></td>
	</tr>


	<tr  align="left" class="itemTD11">
		<td align="left" width="60%"><%=dbRS("Cliente")%></td>
		<td align="left" ><%=dbRS("FechaFactura")%></td>
		<td><%=dbRS("FacturaNumero")%></td>
		<td><%=dbRS("MotivoDescripcion")%></td>
	</tr>
			 
				   
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
<tbody>
	<tr>
		<td colspan="1" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="1" width="50%"><span class="formHeader">Reclamo</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
</tbody>
<tr>
<td width="50%">
<span  class="pagetitle" >
	<span align="left"> ¿Su pedido llegó cerrado?</span>
	<span align="left" class="textFail"><%=dbRS("PedidoCerrado")%> </span>

</span>
	  
<span  class="pagetitle">
	<span>¿Faltó algun producto?</span>
	<span align="left" class="textFail"><%=dbRS("FaltoProducto")%> </span>
	<%if dbRS("FaltoProducto") = "Si" Then %>
	<div align="left">
		<span class="textFail"><%=trim(dbRS("FaltoProductoTexto"))%> </span>
	</div> 
	<%End If%> 
</span> 

<span  class="pagetitle">
	<span>¿Llegó algo en mal estado?</span>
	<span align="left" class="textFail"><%=dbRS("ProductoMalEstado")%> </span>
	<%if dbRS("ProductoMalEstado") = "Si" Then %>
	<div align="left">
		<span class="textFail"><%=trim(dbRS("ProductoMalEstadoTexto"))%> </span>
	</div> 
	<%End If%>	 
</span>

<span class="pagetitle">
	<span>Comentario del cliente</span> 
	<div align="left">
		<span class="textFail"><%=trim(dbRS("Comentario"))%> </span>
	</div> 
</span>
</td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
	<tr>
		<td colspan="1" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="1"><span class="formHeader">Observacion</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
  
	<tr  align="left" class="itemTD11">
		<td align="left" width="60%"><%=dbRS("Observacion")%></td>
 	</tr>
			 
				   
</table>


  