<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC [dbo].[usp_Transferencias_Pendientes_Sel] "  & Session("PUNTO_VENTA_ID") 
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
		<td colspan="4"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="4" align="left"><span class="formHeader">Transferencias Pendientes</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop"> 
		<td height="20" width="20%"><b>Periodo<br/></td> 
		<td height="20" width="20%"><b>Total pendientes<br/></td> 
		<td height="20" width="20%"><b>Importe <br/></td> 
		<td height="20" width="20%"><b>Accion<br/></td> 
	</tr>
  
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF
					
			sSQL =  "EXEC  [usp_Transferencias_Pendientes_view] " & dbRS("RowID") & "," & Session("PUNTO_VENTA_ID") 
			nombrearchivo = "TransferenciasPendientes_" & dbRS("Periodo")  & "_" & dbRS("RowID")
		
		%>
			
	<tr class="itemTD11"> 
		<td><%=dbRS("Periodo")%></td> 
		<td><%=dbRS("TotalPendientes")%></td>
		<td>$ <%=FormatNumber(dbRS("Importe"),2)%></td>  
		<td>
		&nbsp&nbsp&nbsp 
		<a href="javascript:document.FF.doWhat.value='1';chkForm2(document.FF,1,<%=dbRS("RowID")%>)" title="Ver Transfernecias"><img src="../images/eye.png" alt="Ver Transferencias" id="<%=dbRS("RowID")%>"></a>
		&nbsp&nbsp&nbsp 
		<a href="javascript:document.FEXCEL.doWhat.value='';chkFormXls(document.FEXCEL,0,<%=dbRS("RowID")%>,'<%=sSQL%>','<%=nombrearchivo%>')" title="Descargar"><img src="../images/excel.png" alt="Descargar" id="<%=dbRS("RowID")%>"></a>
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
			<td colspan ="4"> No hay transferencias pendientes.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
