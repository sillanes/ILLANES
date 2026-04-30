<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_view "& hojaderuta
'Response.write(sSQL)
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
		<td colspan="11"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="11" align="left"><span class="formHeader">Hojas de Ruta</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="5%"><b>Tipo Factura<br/></td>
		<td height="20" width="5%"><b>Código Fact<br/></td>
		<td height="20" width="5%"><b>Número Factura <br/></td>
		<td height="20" width="5%"><b>Número Cliente <br/></td>
		<td height="20" width="20%"><b>Nombre Cliente<br/></td>
		<td height="20" width="20%"><b>Razón Social<br/></td>
		<td height="20" width="10%"><b>Forma Pago<br/></td>
		<td height="20" width="10%"><b>Importe Facturado<br/></td>
		<td height="20" width="10%"><b>Importe a Cobrar<br/></td>
		<td height="20" width="10%"><b>email<br/></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("RowID")%></td>
		<td><%=dbRS("Tipofact")%></td> 
		<td><%=dbRS("CodigoFact")%></td> 
		<td><%=dbRS("FacturaID")%></td> 
		<td><%=dbRS("ClienteID")%></td>
		<td><%=dbRS("ClienteNombre")%></td>
		<td><%=dbRS("RazonSocial")%></td>
		<td><%=dbRS("FormaPago")%></td> 
		<td><%=dbRS("ImporteFactura")%></td> 
		<td><%=dbRS("ImporteAPagar")%></td> 
		<td><%=dbRS("email")%></td> 
 
		
	</tr>
	
		<%
			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="11"> No hay hojas de ruta en proceso</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
