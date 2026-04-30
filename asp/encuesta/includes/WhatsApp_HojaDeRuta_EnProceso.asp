<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_EnProceso "
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
		<td colspan="7"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="7" align="left"><span class="formHeader">Hojas de Ruta Pendientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="10%"><b>#<br/></td>
		<td height="20" width="30%"><b>Archivo<br/></td>  
		<td height="20" width="10%"><b>clientes</td> 
		<td height="20" width="10%"><b>facturas</td> 
		<td height="20" width="10%"><b>importe</td>
		<td height="20" width="10%"><b>estado</td>  
		<td height="20" width="20%"><b>Accion</td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("HojaDeRutaID")%></td>
		<td><%=dbRS("NombreArchivo")%></td> 
		<td><%=dbRS("Clientes")%></td> 
		<td><%=dbRS("Facturas")%></td>
		<td>$<%=FormatNumber(dbRS("Importe"),2)%></td>     
		<td><%=dbRS("Estado")%></td>    
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("HojaDeRutaID")%>)" title="Ver Hoja de Ruta"><img src="../images/eye.png" alt="Ver Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
		&nbsp&nbsp&nbsp
		<%
		If dbRS("Status")<2 Then 
		%>
		<a href="javascript:document.FF.doWhat.value='2';chkForm2(document.FF,2,<%=dbRS("HojaDeRutaID")%>)"  title="Procesar Hoja de Ruta"><img src="../images/correo.png" alt="Procesar Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:document.FF.doWhat.value='3';chkForm2(document.FF,3,<%=dbRS("HojaDeRutaID")%>)"  title="Eliminar Hoja de Ruta"><img src="../images/del.png" alt="Eliminar Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
		<%
		Else
		%>
		<img src="../images/reloj2.png" alt="Procesando" title="Procesando Hoja de Ruta">
		<%
		End If 
		%>
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
			<td colspan ="7"> No hay Hoja de Rutas en proceso</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
