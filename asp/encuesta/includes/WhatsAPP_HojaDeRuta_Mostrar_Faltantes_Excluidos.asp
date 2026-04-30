<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.[usp_HojaDeRuta_Ver_Faltantes] "& HojaDeRutaID &",1" 
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
				<td colspan="11" align="left"><span class="formHeader">Hoja De Ruta: <%=HojaDeRutaID%> -  EXCLUIDOS</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 

	<tr align="center" class="columnTop" >
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="5%"><b>Cliente<br/></td>
		<td height="20" width="15%"><b>Nombre<br/></td>
		<td height="20" width="15%"><b>Razon Social<br/></td>
		<td height="20" width="5%"><b>Factura<br/></td>
		<td height="20" width="10%"><b>Forma de pago</td>
		<td height="20" width="5%"><b>Importe Factura<br/></td>
		<td height="20" width="5%"><b>Importe a Cobrar<br/></td> 
		<td height="20" width="10%"><b>Telefono<br/></td> 
		<td height="20" width="10%"><b>Factura<br/></td> 
		<td height="20" width="10%"><b>Accion</td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("RowID")%></td>
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
			<a href="javascript:document.FF.doWhat.value='9';chkFormIncluir(document.FF,9,<%=HojaDeRutaID%>,<%=dbRS("RowID")%>)"  title="Incluir Item"><img src="../images/incluir2.png" alt="Excluir Item" id="<%=dbRS("RowID")%>"></a>
			&nbsp&nbsp&nbsp
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
			<td colspan ="11"> No hay datos para la hoja de ruta</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
