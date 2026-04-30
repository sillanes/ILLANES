<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Extractos_Procesados "
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
		<td colspan="12"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="12" align="left"><span class="formHeader">Extractos en Proceso TOP 10</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>#<br/></td>
		<td height="20" width="25%"><b>Archivo<br/></td> 
		<td height="20" width="5%"><b>Banco<br/></td>
		<td height="20" width="5%"><b>Items <br/> Archivo</td>
		<td height="20" width="20%" colspan="5"><b>Cuits<br/></td>
		<td height="20" width="15%"><b>Importe</td> 
		<td height="20" width="5%"><b>Estado</td>  
		<td height="20" width="20%"><b>Accion</td> 
	</tr>
	<tr align="center" class="columnTop">
		<td height="20" width="5%" colspan="4"></td>
		<td height="20" width="8%"><b>Validos<br/></td>
		<td height="20" width="4%"><b>Encontrados<br/></td> 
		<td height="20" width="4%"><b>CuitsOnline<br/></td>  
		<td height="20" width="4%"><b>Manual<br/></td>  
		<td height="20" width="4%"><b>Invalidos<br/></td>  
		<td height="20" width="15%" colspan="3"></td>   
	</tr>
	<tr></tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF
					
			sSQL =  "EXEC  [dbo].[usp_Extractos_download] " & dbRS("ExtractoID")
			nombrearchivo = "Extracto_" & dbRS("BancoNombre")  & "_" & dbRS("ExtractoID")
		
		%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("ExtractoID")%></td>
		<td><%=dbRS("NombreArchivo")%></td> 
		<td><%=dbRS("BancoNombre")%></td> 
		<td><%=dbRS("TotalFilas")%></td> 
		<td><%=dbRS("TotalCuitsArchivo")%></td>  
		<td><%=dbRS("TotalCuitsMatched")%></td> 
		<td><%=dbRS("TotalCuitsOnline")%></td> 
		<td><%=dbRS("TotalCuitsManual")%></td>  
		<td><%=dbRS("TotalCuitsInvalidos")%></td>  
		<td>$ <%=FormatNumber(dbRS("Importe"),2)%></td>   
		<td><%=dbRS("Estado")%></td>  
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("ExtractoID")%>)" title="Ver Extracto"><img src="../images/eye.png" alt="Ver Extracto" id="<%=dbRS("ExtractoID")%>"></a>
		&nbsp&nbsp&nbsp 
		<a href="javascript:document.FEXCEL.doWhat.value='7';chkFormXls(document.FEXCEL,7,<%=dbRS("ExtractoID")%>,'<%=sSQL%>','<%=nombrearchivo%>')" title="Descargar"><img src="../images/excel.png" alt="Descargar" id="<%=dbRS("ExtractoID")%>"></a>
		
		<%If dbRS("Status")>3 and dbRS("Status") <> 5 THEN%>
			&nbsp&nbsp&nbsp 
			<a href="javascript:document.FF.doWhat.value='3';chkForm2(document.FF,3,<%=dbRS("ExtractoID")%>)"  title="Eliminar Extracto"><img src="../images/del.png" alt="Eliminar Extracto" id="<%=dbRS("ExtractoID")%>"></a>
	 		
			<%If dbRS("TotalCuitsOnlineNoEncontrados")>0 or dbRS("TotalCuitsInvalidos") >0  THEN%>
				&nbsp&nbsp&nbsp
				<a href="javascript:document.FF.doWhat.value='4';chkForm2(document.FF,4,<%=dbRS("ExtractoID")%>)" title="Modificar Faltantes"><img src="../images/control.png" alt="Ver Extracto" id="<%=dbRS("ExtractoID")%>"></a>
			<%End IF%>
			
			<%If dbRS("TotalCuitsInvalidos") = 0  THEN%>
				&nbsp&nbsp&nbsp
				<a href="javascript:document.FF.doWhat.value='8';chkForm2(document.FF,8,<%=dbRS("ExtractoID")%>)" title="Aplicar a transferencias"><img src="../images/archivo.png" alt="Aplicar a Transferencias" id="<%=dbRS("ExtractoID")%>"></a>
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
			<td colspan ="11"> No hay extractos en proceso</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
