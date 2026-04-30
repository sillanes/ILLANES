<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_view "& HojaDeRutaID &",1"
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
				<td colspan="11" align="left"><span class="formHeader">Hoja De Ruta: <%=HojaDeRutaID%></span></td>
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
		<td height="20" width="10%"><b>Informacion</td>  
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

		<%If  dbRS("Telefono") <> "" AND dbRS("DirectorioFactura")<>"" THEN%>
			&nbsp&nbsp&nbsp
			<a href="javascript:return false;" title="Completo"><img src="../images/ok.png" alt="Completo"></a>
		<%else%>
			&nbsp&nbsp&nbsp
 			<a href="javascript:return false;" title="Faltan Datos"><img src="../images/warning.png" alt="Faltan Datos"></a>
		<%End IF%>		

		<%If  dbRS("Estado") >= 9 THEN%>
			&nbsp&nbsp&nbsp
			<a href="javascript:return false;" title="Terminado"><img src="../images/candado.png" alt="Terminado"></a>
		<%else%>
			
			
			<%If  dbRS("Enviado") = 1   THEN%>
			
				<%If  dbRS("ResponseCodePDF") =  1   THEN%>
					&nbsp&nbsp&nbsp
					<a href="javascript:return false;" title="Enviado Ok"><img src="../images/enviado2.png" alt="Enviado"></a>
				<%Else%>
					&nbsp&nbsp&nbsp
					<a href="javascript:return false;" title='<%=dbRS("ResponseTextPDF")%>'><img src="../images/enviadoerror.png" alt="Error" ></a>
				<%End IF%>	

			<%Else%>
				<%If  dbRS("ResponseCodePDF") = 1 THEN%>
					&nbsp&nbsp&nbsp
					<a href="javascript:return false;" title="Termiando"><img src="../images/enviado2.png" alt="Terminado"></a>
				<%ElseIf dbRS("ResponseCodePDF") = 2 Then%>
					<a href="javascript:return false;" title='<%=dbRS("ResponseTextPDF")%>'><img src="../images/enviadoerror.png" alt="Error" ></a>
				<%Else%>
					&nbsp&nbsp&nbsp
					<a href="javascript:return false;" title="No enviado"><img src="../images/NoEnviado2.png" alt="No Enviado"></a>	
				<%End IF%>	
			<%End IF%>

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
