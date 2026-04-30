<% 
'Response.write("inicio<br />")
'Response.write(hdrid)
'Response.write(nrocliente)
'Response.write("<br />fin")
sSQL = "EXEC dbo.usp_Transportista_sel "  
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%> 


<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-sm">
	<tbody> 
	<tr>
		<td colspan="4"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="4" align="left"><span class="formHeader">Transportistas</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 

	<tr align="center" class="columnTop" >
		<td height="20" width="5%"><b>Nombre<br/></td>
		<td height="20" width="15%"><b>Apellido<br/></td> 
		<td height="20" width="5%"><b>Telefono<br/></td>
		<td height="20" width="5%"><b>Accion<br/></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
		<%if dbRS("Telefono")<>"" Then%>	
	<tr class="itemTD11">
		<td><%=dbRS("Nombre")%></td> 
		<td><%=dbRS("Apellido")%></td>  
		<td><%=dbRS("Telefono")%></td>    
		<td>
		
		<a href="javascript:document.FF.doWhat.value='2';chkForm2(document.FF,2,<%=HojaDeRutaID%>,<%=dbRS("TransportistaID")%>,0)" title="Selecionar Transportista"><img src="../images/vincular.png" alt="Selecionar Transportista" id="<%=dbRS("TransportistaID")%>"></a>
				
		
		</td>
		
	</tr>
		<%End If%>
		<%
			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="4"> No hay hoja de ruta para visualizar</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
