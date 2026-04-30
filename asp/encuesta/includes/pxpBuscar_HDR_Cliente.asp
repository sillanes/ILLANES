<% 
'Response.write("inicio<br />")
'Response.write(hdrid)
'Response.write(nrocliente)
'Response.write("<br />fin")
sSQL = "EXEC pxp.usp_Buscar_HDR_Cliente " & CLng("0" & hdrid) & "," & CLng("0"&nrocliente)
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%> 
<div style="overflow-x: auto;">
<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable">
	<tbody>
	<tr>
		<td colspan="7" height="20">
		<table border="0" cellpadding="0" cellspacing="0"  width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="7"><span class="formHeader">CLIENTES</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="5" width="5"><b>Hoja De Ruta</b></td>
		<td height="5" width="5"><b>Cliente</b></td>
		<td height="10" width="10"><b>Razon Social</b></td>
		<td height="5" width="10"><b>Total Facturas</b></td>
		<td height="65" width="5"><b>Facturas</b></td>
		<td height="5" width="50"><b>Notas</b></td>
		<td height="5" width="5"><b>Ctrl</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	<tr class="itemTD11">
		<td><%=dbRS("HojaDeRutaNro")%></td>
		<td><%=dbRS("ClienteID")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("RazonSocial")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalFacturas")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Facturas")%>&nbsp;&nbsp;</td>
		<td nowrap>
			<%If  dbRS("Errores")>=0 AND dbRS("Controlado")=1 Then%>
			Armador:<b><%=dbRS("Armador")%></b><br/>
			<b><span <%if dbRS("Errores")>0 then%>style="color:red"<%End If%>>Errores:<%=dbRS("Errores")%></b></span><br/>
			Controlador:<b><%=dbRS("Controlador")%></b><br/>
			Dia:<b><%=dbRS("Dia")%></b><br/>
			Hora:<b><%=dbRS("Hora")%></b><br/>
			<%End If%>
 		</td>		
		<td nowrap>
		<%if dbRS("Controlado")=0 Then %>
		<a href="javascript:document.FF.doWhat.value='2';chkControlar(document.FF,2,<%=dbRS("HojaDeRutaNro")%>, <%=dbRS("ClienteID")%>)"  title="Controlar"><img src="../images/control.png" alt="Controlar" id="<%=dbRS("HojaDeRutaNro")%>"></a>
		<%else%>
		<a title="Controlado"> <img src="../images/candado.png" alt="Controlado" id="<%=dbRS("HojaDeRutaNro")%>"></a>
		<%End If%>
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
			<td colspan ="7"> no se encontraron clientes para la hoja de ruta</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>
</div>


<% 
Set dbRS = Nothing
%>
