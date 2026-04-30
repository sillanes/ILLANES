<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_Reenviados "
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
				<td colspan="11" align="left"><span class="formHeader">Listado reenviados</span></td>
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
		<td height="20" width="5%"><b>Factura<br/></td>
		<td height="20" width="10%"><b>Forma de pago</td>
		<td height="20" width="5%"><b>Importe Factura<br/></td>
		<td height="20" width="5%"><b>Importe a Cobrar<br/></td> 
		<td height="20" width="10%"><b>Telefono<br/></td> 
		<td height="20" width="10%"><b>Factura<br/></td> 
		<td height="20" width="10%"><b>Estado</td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("RowID")%></td>
		<td><%=dbRS("ClienteID")%></td> 
		<td><%=dbRS("ClienteNombre")%></td>  
		<td><%=dbRS("FacturaID")%></td> 
		<td><%=dbRS("FormaPago")%></td> 
		<td>$<%=FormatNumber(dbRS("ImporteFactura"),2)%></td> 
		<td>$<%=FormatNumber(dbRS("ImporteAPagar"),2)%></td> 	 
		<td><%=dbRS("Telefono")%></td> 
		<td><%=dbRS("DirectorioFactura")%></td> 	
		<td>
			<%If  dbRS("Status") = 9 or dbRS("Status") = 10 THEN%>
				<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/candado.png" alt="Terminado">1 de 1 </a>	
				&nbsp&nbsp&nbsp
			<%else%>
				
				<%If  dbRS("Status") < 5 THEN%>
				<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/reloj2.png" alt="Procesando..."></a>	
				&nbsp&nbsp&nbsp
				<%ElseIf  dbRS("Status") = 5 THEN%>
				<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/reloj2.png" alt="Procesando..."></a>	
				&nbsp&nbsp&nbsp
				<%ElseIf dbRS("Status") = 6 Then  %>
				<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/FTPupload.png" alt="Subiendo archivos..."></a>	
				&nbsp&nbsp&nbsp
				<%ElseIf  dbRS("Status") = 7 Then %>
				<a href="javascript:return false;" title=<%=dbRS("Estado")%>"><img src="../images/FTPuploadOk.png" alt="Procesando..."></a>	
				&nbsp&nbsp&nbsp
				<%ElseIf  dbRS("Status") = 8 Then %>
				<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/reloj2.png" alt="Procesando..."> <%=dbRS("Enviados")%> de <%=dbRS("Clientes")%></a>	
				&nbsp&nbsp&nbsp
				<%End If%>
			
			<%End IF%>
			
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
			<td colspan ="11"> No hay hoja de ruta para visualizar</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>