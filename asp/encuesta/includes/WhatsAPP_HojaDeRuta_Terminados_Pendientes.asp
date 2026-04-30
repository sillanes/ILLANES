<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_HojaDeRuta_Terminados_Pendientes "
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
		<td colspan="13"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="13" align="left"><span class="formHeader">Hojas de ruta en proceso</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"></td>
		<td height="20" width="15%"></td>  
		<td height="20" width="5%" colspan="3" style="border-style: solid;border-color: #96D4D4;"><b>Clientes</b></td> 
		<td height="20" width="5%" colspan="3" style="border-style: solid;;border-color: #FF0000;"><b>Envios</b></td> 
		<td height="20" width="5%"></td>   
		<td height="20" width="15%"></td> 
	</tr> 	
	<tr align="center"  class="columnTop">
		<td height="20" width="5%"><b>#<br/></b></td>
		<td height="20" width="15%"><b>Archivo</b></td>  
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Totales</td>
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Validos</td> 
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Excluidos</td>
		<td height="20" width="5%" style="border-style: solid;;border-color: #FF0000;">Envios</td>
		<td height="20" width="5%" style="border-style: solid;;border-color: #FF0000;">Ok</td> 
		<td height="20" width="5%" style="border-style: solid;;border-color: #FF0000;">Con Error</td>   
		<td height="20" width="5%"><b>Estado</b></td>  
		<td height="20" width="15%"><b>Accion</b></td> 
	</tr> 
	<tr>
	<td colspan="10"></td>
	</tr> 
	
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF
					
			sSQL =  "EXEC  [dbo].[usp_HojaDeRuta_Terminados_Pendientes_download] " & dbRS("HojaDeRutaID")
			nombrearchivo = "HojaDeRuta_" & dbRS("HojaDeRutaID")
		
		%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("HojaDeRutaID")%></td>
		<td><%=dbRS("NombreArchivo")%></td>  
		<td><%=dbRS("ClientesTotales")%></td> 
		<td><%=dbRS("ClientesIncluidos")%></td>
		<td><%=dbRS("ClientesExcluidos")%></td>  
		<td><%=dbRS("EnviosTotales")%></td> 
		<td><%=dbRS("EnviosOk")%></td>  
		<td><%=dbRS("EnviosError")%></td>  
		<td><%=dbRS("Estado")%></td> 
 
	  
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("HojaDeRutaID")%>)" title="Ver Hoja de Ruta"><img src="../images/eye.png" alt="Ver Hoja de Ruta" id="<%=dbRS("HojaDeRutaID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:document.FEXCEL.doWhat.value='7';chkFormXls(document.FEXCEL,7,<%=dbRS("HojaDeRutaID")%>,'<%=sSQL%>','<%=nombrearchivo%>')" title="Descargar"><img src="../images/excel.png" alt="Descargar" id="<%=dbRS("HojaDeRutaID")%>"></a>
		&nbsp&nbsp&nbsp
		<a href="javascript:return false;" title="<%=dbRS("Estado")%>"><img src="../images/candado.png" alt="Terminado"> <%=dbRS("EnviosOK")%> de <%=dbRS("EnviosTotales")%></a>	
		&nbsp&nbsp&nbsp
		   
 
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
			<td colspan ="10"> No hay hojas de ruta en proceso</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
