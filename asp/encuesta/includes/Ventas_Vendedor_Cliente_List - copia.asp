<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC dbo.Vendedor_Clientes_List " &"'"& VendedorID & "'," & PeriodoID
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
<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="8" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="8"><span class="formHeader">Clientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Cliente Nro</b></td>
		<td height="20" width="30"><b>Nombre<br/>Cliente</b></td> 
		<td height="20" width="30"><b>Razon<br/>Social</b></td>
		<td height="20" width="30"><b>Total<br/>Objetivos </b></td>
		<td height="20" width="30"><b>Cantidad Fotos Subidas</b></td>
		<td height="20" width="30"><b>Cantidad Fotos Validas</b></td>
		<td height="40" width="50"><b>Cumplimiento</b></td> 
		<td height="20" width="30"><b>Accion</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("ClienteID")%></td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("ClienteNombre")%>&nbsp;&nbsp;</td>
		<td nowrap>&nbsp;&nbsp;<%=dbRS("RazonSocial")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalObjetivos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalSubidas")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalValido")%>&nbsp;&nbsp;</td>	
		<td nowrap> 
		
		<progress-meter>
          <progress-percent style="--progress: <%=dbRS("CumplimientoValidadas")%>"></progress-percent>
        </progress-meter>
		
		
		</td>
 
		<td><a href="javascript:chkForm2(document.FF2,1,'<%=dbRS("ClienteID")%>','<%=dbRS("ClienteNombre")%>')" title="Ver Objetivos"><img src="../images/eye.png" alt="Ver Reclamo" id="<%=dbRS("ClienteID")%>"></a></td>
	</tr>
	
		<%
	 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="8"> No tiene clientes asociados</a>.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>
</div>

<% 
Set dbRS = Nothing
%>
