<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Transportista_HojaDeRuta_SinAsignar_Sel "
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
		<td colspan="5"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="5" align="left"><span class="formHeader">Hojas de Ruta Pendientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="10%"><b>#<br/></td> 
		<td height="20" width="10%"><b>clientes</td> 
		<td height="20" width="10%"><b>facturas</td>  
		<td height="20" width="20%"><b>Accion</td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("HojaDeRutaID")%></td>  
		<td><%=dbRS("Clientes")%></td> 
		<td><%=dbRS("Facturas")%></td>  
		<td>
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("HojaDeRutaID")%>,0,0)" title="Asignar Transportista"><img src="../images/vincular.png" alt="Asignar Transportista" id="<%=dbRS("HojaDeRutaID")%>"></a>
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
			<td colspan ="5"> No hay Hoja de Rutas en proceso</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
