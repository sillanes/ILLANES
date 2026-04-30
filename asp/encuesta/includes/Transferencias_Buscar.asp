<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Transferencias_Buscar "  & ipendientes & ","& idMonto &",'" & idCliente &"'," & irango & "," & txtStartRange & "," & txtEndRange& ",''," & Session("PUNTO_VENTA_ID")  
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
		<td colspan="9"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="9" align="left"><span class="formHeader">Transferencias</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop" > 
		<td height="20" width="5%"><b>Fecha<br/></td>
		<td height="20" width="5%"><b>Banco<br/></td>
		<td height="20" width="5%"><b>Cliente<br/></td>
		<td height="20" width="5%"><b>Monto<br/></td>
		<td height="20" width="5%"><b>F. Deposito<br/>ECHEQ</td> 
		<td height="20" width="5%"><b>OK</td>  
		<td height="20" width="5%"><b>NRO Recibo</td>  
		<td height="20" width="5%"><b>NRO Cliente</td>  
		<td height="20" width="5%"><b>Acción</td>  
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11"> 
		<td><%=dbRS("Fecha")%></td>
		<td><%=dbRS("Banco")%></td> 
		<td><%=dbRS("Cliente")%></td> 
		<td>$ <%=FormatNumber(dbRS("Monto"),2)%></td> 
		<td><%=dbRS("FechaDepositoECheck")%></td> 
		<td><%=dbRS("OK")%></td> 
		<td><%=dbRS("NroRecibo")%></td> 
		<td><%=dbRS("NroCliente")%></td>  
		<td width="140px">
		<a href="javascript:document.FF.doWhat.value='2';chkForm3(document.FF,2,<%=dbRS("RowID")%>,<%=idMonto%>)" title="Modificar"><img src="../images/control.png" alt="Registrar" id="<%=dbRS("RowID")%>"></a>
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
			<td colspan ="9"> No se encontraron transferencias para las paramentros seleccionados</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
