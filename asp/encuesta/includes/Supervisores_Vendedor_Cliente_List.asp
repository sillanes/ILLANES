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
<table  border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable">
	<tbody> 
	<tr>
		<td colspan="9" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="9"><span class="formHeader">Clientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop"> 
		<td height="20" width="30"><b>Accion</b></td>
		<td height="40" width="50"><b>Cumplimiento</b></td> 
		<td height="20" width="30"><b>Cliente</b></td>  
		<td height="20" width="30"><b>Total<br/>Objetivos </b></td>
		<td height="20" width="30"><b>Fotos Subidas</b></td>
		<td height="20" width="30"><b>Fotos Validas</b></td>
		<td height="20" width="30"><b>Fotos Pendientes</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">

		<td><a href="javascript:chkForm2(document.FF3,2,'<%=dbRS("ClienteID")%>','<%=dbRS("ClienteNombre")%>')" title="Ver Objetivos"><img src="../images/eye.png" alt="Ver Objetivos" id="<%=dbRS("ClienteID")%>"></a></td>
		<td nowrap> 
		<progress-meter>
          <progress-percent style="--progress: <%=dbRS("CumplimientoValidadas")%>"></progress-percent>
        </progress-meter> 
		</td>
 
		<%if Session("isMobile") Then%> 
		<td nowrap onclick="ShowComments('<%=dbRS("ClienteNombre")%>')"><%=dbRS("ClienteID")%></td>
		<%else%>
		<td><%=dbRS("ClienteNombre")%></td>
		<%End If%>

 
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalObjetivos")%>&nbsp;&nbsp;</td>	
		<td nowrap>&nbsp;&nbsp;<%=dbRS("TotalSubidas")%>&nbsp;&nbsp;</td>	
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
			<td colspan ="9"> No tiene clientes asociados</a>.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>
</div>

<% 
Set dbRS = Nothing
%>
