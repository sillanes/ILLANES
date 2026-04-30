<%

sSQL = "EXEC usp_Facturas_Cliente_Sel " & CLng("0"&idCliente) & "," & CLng("0"&dbStartDate)
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>
<input type="hidden" name="email" value="<%=email%>">
<input type="hidden" name="telefono" value="<%=telefono%>">

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%">
	<tbody>
	<tr>
		<td colspan="3" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="3"><span class="formHeader">FACTURAS</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Nro Factura</b></td>
		<td height="20" width="30"><b>Tipo Factura</b></td>
		<td height="20" width="30"><b>Importe</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	<tr class="itemTD11" onclick="chkForm2(document.FF, 3, <%=dbRS("FacturaID")%>)">
		<td><%=dbRS("FacturaID")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("CodFact")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalFactura")%>&nbsp;&nbsp;</td>	
	</tr>
	
		<%
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="3"> No se encuentran facturas para la fecha seleccionada, verifique selección o comuniquese con <a href="mailto:reclamos@illanes.com.ar">reclamos@illanes.com.ar</a>.</td>
		</tr>
		<tr class="itemTD11"><td align="center" colspan ="3"> <input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)"></td></tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
