<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Extractos_view "& Extracto
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
				<td colspan="7" align="left"><span class="formHeader">Extracto: <%=Extracto%></span></td>
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
		<td height="20" width="5%"><b>Nombre en Archivo<br/></td>
		<td height="20" width="5%"><b>CUIT Archivo<br/></td>
		<td height="20" width="5%"><b>CUIT Validado<br/></td> 
		<td height="20" width="5%"><b>Razon Social<br/>Validado</td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("RowID")%></td>
		<td><%=dbRS("Fecha")%></td> 
		<td>$<%=FormatNumber(dbRS("Credito"),2)%></td> 
		<td><%=dbRS("NombreArchivo")%></td> 
		<td><%=dbRS("CUIT_Archivo")%></td> 
		<td><%=dbRS("CUIT_Valido")%></td> 
		<td><%=dbRS("RazonSocial")%></td>  
 
		
	</tr>
	
		<%
			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="7"> No hay extracto para visualizar</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
