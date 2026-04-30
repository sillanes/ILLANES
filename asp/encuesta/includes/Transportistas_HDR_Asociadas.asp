<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Transportista_HojaDeRuta_Asociadas_Sel "
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

Dim cols
If Session("currentUser") <> "logistica" Then
	cols =8
Else
	cols =7
End If

   
%>
<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody> 
	<tr>
		<td colspan="<%=cols%>"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="8" align="left"><span class="formHeader">Hojas de Ruta</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>#<br/></td> 
		<td height="20" width="5%"><b>Telefono</td> 
		<td height="20" width="15%"><b>Repartidor 1</td>  
		<td height="20" width="15%"><b>Repartidor 2</td>  
		<td height="20" width="1%"><b>Patente</td>  
		<td height="20" width="1%"><b>Contraseña</td> 
		<td height="20" width="70%"><b>Link</td> 
		
<% if Session("currentUser") <> "logistica" Then %>
		<td height="20" width="70%"><b>Accion</td> 
<% End If %>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("HojaDeRutaID")%></td>  
		<td><%=dbRS("Telefono")%></td> 
		<td><%=dbRS("Repartidor1")%></td>  
		<td><%=dbRS("Repartidor2")%></td>  
		<td><%=dbRS("Patente")%></td>   
		<td><%=dbRS("pass")%></td>   
		<td><%=dbRS("Link")%></td>
<% if Session("currentUser") <> "logistica" Then %> 
		<td><a href="javascript:document.FF.doWhat.value='5';chkForm2(document.FF,5,<%=dbRS("HojaDeRutaID")%>,0,0)" title="Eliminar HDR"><img src="../images/del.png" alt="Asignar Transportista" id="<%=dbRS("HojaDeRutaID")%>"></a>
 <% End If %>
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
			<td colspan ="<%=cols%>"> No hay Hoja de Rutas asociadas</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
