<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC super.Supervisor_Vendedores_List " &"'"& SupervisorID & "'," & PeriodoID
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
<table  border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable">
	<tbody> 
	<tr>
		<td colspan="10" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" >
			<tbody><tr class="tableHeader">
				<td align="left" colspan="10"><span class="formHeader">Clientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">  
		<td height="40" width="50"><b>Cumplimiento</b></td> 
		<td height="20" width="30"><b>Vendedor</b></td>  
		<td height="20" width="30"><b>Total<br/>Clientes </b></td>
		<td height="20" width="30"><b>Total<br/>Objetivos</b></td>
		<td height="20" width="30"><b>Fotos Subidas</b></td>
		<td height="20" width="30"><b>Fotos Descartadas</b></td>
		<td height="20" width="30"><b>Fotos Validas</b></td>
		<td height="20" width="30"><b>Fotos Pendientes</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
 
		<td nowrap> 
		
		<progress-meter>
          <progress-percent style="--progress: <%=dbRS("CumplimientoValidadas")%>"></progress-percent>
        </progress-meter>
		
		
		</td>
 
		<%if Session("isMobile") Then%>
 
		<td nowrap onclick="ShowComments('<%=dbRS("Nombre")%>')"><%=dbRS("VendedorID")%></td>
		<%else%>
		<td><%=dbRS("Nombre")%></td>
		<%End If%>

 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalClientes")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalObjetivos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("Subidas")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("SubidasDescartadas")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalValido")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalPendientes")%>&nbsp;&nbsp;</td>
	</tr>
	
		<%
	 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="10"> No tiene vendedores asociados</a>.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>
</div>

<% 
Set dbRS = Nothing
%>
