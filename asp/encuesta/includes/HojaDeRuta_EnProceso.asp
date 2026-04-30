<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_Sel "
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
		<td colspan="10"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="10" align="left"><span class="formHeader">Hojas de Ruta en proceso</span></td>
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
		<td height="20" width="8%"><b>Emails Tokin<br/></td>
		<td height="20" width="8%"><b>Emails Reclamos<br/></td>
		<td height="20" width="8%"><b>Importe Facturado<br/></td>
		<td height="20" width="8%"><b>Importe a Cobrar<br/></td> 
		<td height="20" width="8%"><b>estado<br/></td>
		<td height="20" width="8%"><b>Total Emails Enviados<br/></td>
		<td height="20" width="50%"><b>Acción<br/></td>
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
		<td><%=dbRS("estado")%></td> 
		<td><%=dbRS("TotalEmailEnviados")%></td> 
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("RowID")%>)" title="Ver Hoja de Ruta"><img src="../images/eye.png" alt="Ver Hoja de Ruta" id="<%=dbRS("RowID")%>"></a>
		&nbsp&nbsp&nbsp
		<%
		If dbRS("Status")<2 Then 
		%>
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,2,<%=dbRS("RowID")%>)"  title="Enviar Hoja de Ruta"><img src="../images/correo.png" alt="Enviar Hoja de Ruta" id="<%=dbRS("RowID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,3,<%=dbRS("RowID")%>)"  title="Eliminar Hoja de Ruta"><img src="../images/del.png" alt="Eliminar Hoja de Ruta" id="<%=dbRS("RowID")%>"></a>
		<%
		Else
		%>
		<img src="../images/reloj2.png" alt="Procesando" title="Enviando hoja de ruta">
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
			<td colspan ="10"> No hay hojas de ruta en proceso</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
