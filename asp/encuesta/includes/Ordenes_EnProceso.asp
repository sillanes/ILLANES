<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Ordenes_EnProceso "
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
				<td colspan="8" align="left"><span class="formHeader">Ordenes Pendientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="20%"><b>Archivo<br/></td>
		<td height="20" width="10%"><b>Total Items Archivo<br/></td> 
		<td height="20" width="4%"><b>Cuits Macheados<br/></td> 
		<td height="20" width="4%"><b>Cuits No encontrados<br/></td>  
		<td height="20" width="8%"><b>Vendedores No Encontrados<br/></td> 
		<td height="20" width="8%"><b>Estado<br/></td> 
		<td height="20" width="10%"><b>Accion<br/></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("FileID")%></td>
		<td><%=dbRS("NombreArchivo")%></td> 
		<td><%=dbRS("TotalClientes")%></td> 
		<td><%=dbRS("TotalClientesMatched")%></td> 
		<td><%=dbRS("ClientesMissing")%></td>    
		<td><%=dbRS("VendedoresMissing")%></td>   
		<td><%=dbRS("Estado")%></td>  
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("FileID")%>)" title="Ver Orden"><img src="../images/eye.png" alt="Ver Orden" id="<%=dbRS("FileID")%>"></a>
		&nbsp&nbsp&nbsp
		<%
		If dbRS("Status")<2 Then 
		%>
		<a href="javascript:document.FF.doWhat.value='2';chkForm2(document.FF,2,<%=dbRS("FileID")%>)"  title="Procesar Orden"><img src="../images/correo.png" alt="Procesaro Orden" id="<%=dbRS("FileID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:document.FF.doWhat.value='3';chkForm2(document.FF,3,<%=dbRS("FileID")%>)"  title="Eliminar Orden"><img src="../images/del.png" alt="Eliminar Orden" id="<%=dbRS("FileID")%>"></a>
		<%
		Else
		%>
		<img src="../images/reloj2.png" alt="Procesando" title="Procesando Ordenes">
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
			<td colspan ="8"> No hay Ordenes en proceso</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
