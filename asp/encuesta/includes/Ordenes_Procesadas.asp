<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Ordenes_Procesadas "
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
		<td colspan="15"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="15" align="left"><span class="formHeader">Files en Proceso TOP 10</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	
	<tr align="center" class="columnTop">
		<td height="20" width="5%" rowspan="2">#</td>
		<td height="20" width="8%" rowspan="2"><b>Archivo</b><br/></td>
		<td height="20" width="4%" rowspan="2"><b>Clientes Totales</b><br/></td> 
		<td height="20" width="4%" rowspan="2"><b>Clientes No Encontrados</b><br/></td>  
		<td height="20" width="4%" rowspan="2"><b>Vendedores No Encontrados</b><br/></td>
		<td height="20" width="4%" colspan="6"><b>Cajas</b><br/></td>    
		<td height="20" width="4%" rowspan="2"><b>Estado</b><br/></td>      
		<td height="20" width="6%" rowspan="2"><b>Pendientes</b><br/></td>       
		<td height="20" width="6%" rowspan="2"><b>Verificado</b><br/></td>    
		<td height="20" width="6%" rowspan="2"><b>Accion</b></td>   
	</tr>  
	<tr align="center" class="columnTop">    
    <td height="20" width="4%"><b>Bronce</b></td>      
    <td height="20" width="4%"><b>Plata</b></td>   
    <td height="20" width="4%"><b>Oro</b></td>    
    <td height="20" width="4%"><b>Platino</b></td>   
    <td height="20" width="4%"><b>Especial</b></td>   
    <td height="20" width="4%"><b>Estrella</b></td>    
	</tr> 
	<tr>
	<td colspan="12"></td>
	</tr> 
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF
					
			sSQL =  "EXEC  [dbo].[usp_Ordenes_download] " & dbRS("FileID")
			nombrearchivo = "Ordenes_Empretienda_" & dbRS("FileID")
		
		%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("FileID")%></td>
		<td><%=dbRS("NombreArchivo")%></td>  
		<td><%=dbRS("TotalClientes")%></td> 
		<td><%=dbRS("ClientesMissing")%></td>  
		<td><%=dbRS("VendedoresMissing")%></td>    
		<td><%=dbRS("CajasBronce")%></td>
		<td><%=dbRS("CajasPlata")%></td>
		<td><%=dbRS("CajasOro")%></td>  
		<td><%=dbRS("CajasPlatino")%></td>   
		<td><%=dbRS("CajasEspecial")%></td>  
		<td><%=dbRS("CajasEstrella")%></td>    
		<td><%=dbRS("Estado")%></td>  
		<td>
		<%If  dbRS("ClientesMissing") = 0  THEN%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:return false;" title="Clientes Completo"><img src="../images/ok.png" alt="Clientes Completo"></a>
		<%else%>
			&nbsp;&nbsp;&nbsp;
 			<a href="javascript:return false;" title="Faltan Datos Clientes"><img src="../images/warning.png" alt="Faltan Datos Clientes"></a>
		<%End IF%>  
		
		<%If dbRS("VendedoresMissing") = 0 THEN%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:return false;" title="Vendedores Completo"><img src="../images/ok.png" alt="Vendedores Completo"></a>
		<%else%>
			&nbsp;&nbsp;&nbsp;
 			<a href="javascript:return false;" title="Faltan Datos Vendedores"><img src="../images/warning.png" alt="Faltan Datos Vendedores"></a>
		<%End IF%>	
		
		<%If dbRS("FacturasMissing") = 0 AND dbRS("FacturaArchivoMissing") = 0  THEN%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:return false;" title="Facturas Completo"><img src="../images/ok.png" alt="Facturas Completo"></a>
		<%else%>
			&nbsp;&nbsp;&nbsp;
 			<a href="javascript:return false;" title="Faltan Datos Factura"><img src="../images/warning.png" alt="Faltan Datos Factura"></a>
		<%End IF%>			
		</td> 

		<td>
			<%If dbRS("FacturaValidadaMissing") = 0 AND dbRS("FacturaValidadaWarning") = 0   THEN%>
				&nbsp;&nbsp;&nbsp;
				<a href="javascript:return false;" title="Facturas Vaidadas"><img src="../images/ok.png" alt="Facturas Vaidadas"></a>
			<%else%>
				&nbsp;&nbsp;&nbsp;
				<a href="javascript:return false;" title="Pendintes de validacion"><img src="../images/warning.png" alt="Faltan Datos Factura"></a>
			<%End IF%>	
		</td>		
		
		<td>
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("FileID")%>)" title="Ver Ordenes"><img src="../images/eye.png" alt="Ver Ordenes" id="<%=dbRS("FileID")%>"></a>
		&nbsp&nbsp&nbsp 
		<a href="javascript:document.FEXCEL.doWhat.value='7';chkFormXls(document.FEXCEL,7,<%=dbRS("FileID")%>,'<%=sSQL%>','<%=nombrearchivo%>')" title="Descargar"><img src="../images/excel.png" alt="Descargar" id="<%=dbRS("FileID")%>"></a>
		
		<%If dbRS("Status")>3 and dbRS("Status") < 5 THEN%>
			&nbsp&nbsp&nbsp 
			<a href="javascript:document.FF.doWhat.value='3';chkForm2(document.FF,3,<%=dbRS("FileID")%>)"  title="Eliminar Orden"><img src="../images/del.png" alt="Eliminar Orden" id="<%=dbRS("FileID")%>"></a>
	 		
			<%If dbRS("ClientesMissing")>0 or dbRS("VendedoresMissing") >0  or dbRS("FacturaArchivoMissing")>0 THEN%>
				&nbsp&nbsp&nbsp
				<a href="javascript:document.FF.doWhat.value='4';chkForm2(document.FF,4,<%=dbRS("FileID")%>)" title="Modificar Faltantes"><img src="../images/control.png" alt="Ver Orden" id="<%=dbRS("FileID")%>"></a>
			<%End IF%>
			
			<%If dbRS("ClientesMissing")=0 AND dbRS("VendedoresMissing") = 0 AND dbRS("FacturasMissing") = 0 AND dbRS("FacturaArchivoMissing") = 0 THEN%>
				&nbsp&nbsp&nbsp
				<a href="javascript:document.FF.doWhat.value='8';chkForm2(document.FF,8,<%=dbRS("FileID")%>)" title="Guardar"><img src="../images/archivo.png" alt="Guardar" id="<%=dbRS("FileID")%>"></a>
			<%End IF%>
				 
		<%End IF%>
		
		<%If dbRS("Status")=5 THEN%>
				&nbsp&nbsp&nbsp 
				<a href="javascript:document.FF.doWhat.value='10';chkForm2(document.FF,10,<%=dbRS("FileID")%>)" title="Enviar Email"><img src="../images/email.png" alt="enviar email" id="<%=dbRS("FileID")%>"></a>
		<%End IF%>
	
		<%If dbRS("Status")=6 THEN%>
				&nbsp&nbsp&nbsp 
				<a href="javascript:return false;" title="Enviando..."><img src="../images/reloj2.png" alt="Enviando..." id="<%=dbRS("FileID")%>"></a>
		<%End IF%>

		<%If dbRS("Status")=7 THEN%>
				&nbsp&nbsp&nbsp 
				<a href="javascript:return false;" title="Enviado"><img src="../images/enviado2.png" alt="enviar email" id="<%=dbRS("FileID")%>"></a>
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
			<td colspan ="13"> No hay ordenes en proceso</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
