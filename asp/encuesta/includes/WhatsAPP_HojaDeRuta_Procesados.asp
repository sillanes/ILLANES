<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_Procesados "
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
		<td colspan="13"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="13" align="left"><span class="formHeader">Hojas de ruta en proceso</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="15%"><b>Archivo<br/></td>  
		<td height="20" width="5%"><b>clientes</td> 
		<td height="20" width="5%"><b>facturas</td> 
		<td height="20" width="5%"><b>importe</td>
		<td height="20" width="5%"><b>Total <br/> Telefonos</td> 
		<td height="20" width="5%"><b>Total <br/> Archivos Factura</td>  
		<td height="20" width="5%"><b>Total <br/> Subidos a FTP</td>
		<td height="20" width="5%"><b>Total <br/> Enviados</td>
		<td height="20" width="10%"><b>estado<br/>general</td> 
		<td height="20" width="5%"><b>estado<br/>Telefonos</td>   
		<td height="20" width="5%"><b>estado<br/>PDFs</td>  
		<td height="20" width="15%"><b>Accion</td> 
	</tr> 
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF
					
			sSQL =  "EXEC  [dbo].[usp_HojaDeRuta_download] " & dbRS("HojaDeRutaID")
			nombrearchivo = "HojaDeRuta_" & dbRS("HojaDeRutaID")
		
		%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("HojaDeRutaID")%></td>
		<td><%=dbRS("NombreArchivo")%></td>  
		<td><%=dbRS("Clientes")%></td> 
		<td><%=dbRS("Facturas")%></td>  
		<td>$ <%=FormatNumber(dbRS("Importe"),2)%></td>   
		<td><%=dbRS("TotalTelefonos")%></td> 
		<td><%=dbRS("TotalDirectorios")%></td>  
		<td><%=dbRS("SubidosFTP")%></td>  
		<td><%=dbRS("Enviados")%></td> 
		<td>
			
			<%=dbRS("Estado")%>
			<%If  dbRS("Status") = 9 THEN%>
				<img src="../images/enviado2.png" alt="Enviado">
			<%End If%>
			
			<%If dbRS("Status") = 10 THEN%>
				<img src="../images/enviadoerror.png" alt="Enviado con errores" >
			<%End If%>
						
		</td>   
		<td>
			<%If dbRS("EstadoTelefonos") = 1 Then%>
				<img src="../images/ok.png" alt="Completo">
			<%Else%>
				<img src="../images/warning2.png" alt="Incompleto">
			<%End If%>
		</td> 
		<td> 
			<%If dbRS("EstadoPdfs") = 1 Then%>
				<img src="../images/ok.png" alt="Completo">
			<%Else%>
				<img src="../images/warning2.png" alt="Incompleto">
			<%End If%>
		</td>  
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("HojaDeRutaID")%>)" title="Ver Hoja de Ruta"><img src="../images/eye.png" alt="Ver Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:document.FEXCEL.doWhat.value='7';chkFormXls(document.FEXCEL,7,<%=dbRS("HojaDeRutaID")%>,'<%=sSQL%>','<%=nombrearchivo%>')" title="Descargar"><img src="../images/excel.png" alt="Descargar" id="<%=dbRS("HojaDeRutaID")%>"></a>
		&nbsp&nbsp&nbsp 
		<%If  dbRS("Status") < 5  THEN%>
			
			<a href="javascript:document.FF.doWhat.value='3';chkForm2(document.FF,3,<%=dbRS("HojaDeRutaID")%>)"  title="Borrar hoja de ruta"><img src="../images/del.png" alt="Borrar Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
			&nbsp&nbsp&nbsp
		
			<%If dbRS("Facturas")<>dbRS("TotalTelefonos") or dbRS("TotalDirectorios") <> dbRS("Facturas")  THEN%>
				<a href="javascript:document.FF.doWhat.value='4';chkForm2(document.FF,4,<%=dbRS("HojaDeRutaID")%>)" title="Modificar Faltantes"><img src="../images/control.png" alt="Ver Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
				&nbsp&nbsp&nbsp
			<%End IF%>
			
			<%If dbRS("Facturas")=dbRS("TotalTelefonos") AND dbRS("Facturas") = dbRS("TotalDirectorios") THEN%>
				<a href="javascript:document.FF.doWhat.value='4';chkForm2(document.FF,4,<%=dbRS("HojaDeRutaID")%>)" title="Modificar Faltantes"><img src="../images/control.png" alt="Ver Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
				&nbsp&nbsp&nbsp
			<%End IF%>
			
			<%If (dbRS("Facturas")=dbRS("TotalTelefonos") AND dbRS("Facturas") = dbRS("TotalDirectorios")) OR dbRs("FacturaObligatoria") = 0  THEN%>
				<a href="javascript:document.FF.doWhat.value='8';chkForm2(document.FF,8,<%=dbRS("HojaDeRutaID")%>)" title="Enviar WhatsApp"><img src="../images/whastapplogo2.png" alt="Enviar a WhatsApp" id="<%=dbRS("HojaDeRutaID")%>"></a>
				&nbsp&nbsp&nbsp
			<%End IF%>			
				 
		<%Else%>
		
			<%If  dbRS("Status") = 9 or dbRS("Status") = 10 THEN%>
				<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/candado.png" alt="Terminado"> <%=dbRS("Enviados")%> de <%=dbRS("Clientes")%></a>	
				&nbsp&nbsp&nbsp
			<%else%>
				
				<%If  dbRS("Status") = 5 THEN%>
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
			<td colspan ="13"> No hay hojas de ruta en proceso</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
