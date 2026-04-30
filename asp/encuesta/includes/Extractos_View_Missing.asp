<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Extractos_view_missing "& Extracto
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
				<td colspan="8" align="left"><span class="formHeader">Datos</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop" >
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="5%"><b>Fecha<br/></td>
		<td height="20" width="5%"><b>Credito<br/></td>
		<td height="20" width="5%"><b>Descripcion<br/></td>
		<td height="20" width="5%"><b>CUIT Archivo<br/></td>
		<td height="20" width="5%"><b>CUIT Validado<br/></td> 
		<td height="20" width="5%"><b>Razon Social<br/>Validado</td>  
		<td height="20" width="5%"><b>Accion</td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("RowID")%></td>
		<td><%=dbRS("Fecha")%></td> 
		<td><%=dbRS("Credito")%></td> 
		<td><%=dbRS("NombreArchivo")%></td> 
		<td><%=dbRS("CUIT_Archivo")%></td> 
		<td><%=dbRS("CUIT_Valido")%></td> 
		<td><%=dbRS("RazonSocial")%></td> 
		<td>
 		<a href="javascript:document.FF.doWhat.value='5';chkForm3(document.FF,5,<%=Extracto%>,<%=dbRS("RowID")%>)"  title="Modificar CUIT"><img src="../images/control.png" alt="Modificar CUIT" id="<%=dbRS("RowID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:document.FF.doWhat.value='11';chkForm3(document.FF,11,<%=Extracto%>,<%=dbRS("RowID")%>)"  title="Eliminar Extracto"><img src="../images/del.png" alt="Eliminar Extracto" id="<%=dbRS("RowID")%>"></a>
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
			<td colspan ="8"> No hay extracto para visualizar</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
