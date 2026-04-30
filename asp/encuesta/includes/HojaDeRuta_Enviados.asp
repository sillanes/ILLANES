<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_Enviadas "
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
		<td colspan="8"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="8" align="left"><span class="formHeader">Hojas de Ruta enviadas</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="20%"><b>Archivo<br/></td>
		<td height="20" width="10%"><b>Clientes Totales<br/></td>
		<td height="20" width="5%"><b>Emails Tokin<br/></td>
		<td height="20" width="10%"><b>Emails Reclamos<br/></td>
		<td height="20" width="10%"><b>Importe Facturado<br/></td>
		<td height="20" width="10%"><b>Importe a Cobrar<br/></td>		
		<td height="20" width="10%"><b>Total Enviados<br/></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("RowID")%></td>
		<td><%=dbRS("NombreArchivo")%></td> 
		<td><%=dbRS("Clientes")%></td> 
		<td><%=dbRS("ClientesTokin")%></td> 
		<td><%=dbRS("ClientesEmailReclamo")%></td> 
		<td><%=dbRS("ImporteFacturado")%></td> 
		<td><%=dbRS("ImporteaCobrar")%></td> 		
		<td><%=dbRS("TotalEmailEnviados")%></td> 
	</tr>
	
		<%
			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="8"> No hay hojas de ruta en enviadas</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
