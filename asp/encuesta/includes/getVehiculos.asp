<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC dbo.usp_Vehiculos_sel 0"
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>

<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-xl">
	<tbody> 
	<tr>
		<td colspan="3" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="3"><span class="formHeader">Nomina de vehiculos</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="30"><b>Patente</td>
		<td height="20" width="30"><b>Estado</td>
		<td height="20" width="30"><b>Acción</b></td> 
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td><%=dbRS("Patente")%></td>   
		<td>
		<%If  dbRS("activo") = 1 THEN%>
			&nbsp&nbsp&nbsp
			<a href="javascript:return false;" title="Activo"><img src="../images/ok.png" alt="Completo"></a>
		<%else%>
			&nbsp&nbsp&nbsp
 			<a href="javascript:return false;" title="Inactivo"><img src="../images/candado.png" alt="inactivo"></a>
		<%End IF%>	
		</td>
		<td>
			<a href="javascript:document.FF.doWhat.value='2';chkForm2(document.FF,2,'<%=dbRS("patente")%>')" title="Desactivar"><img src="../images/del.png" alt="Desactivar"></a>
			
			<%If  dbRS("activo") =0 THEN%>
			<a href="javascript:document.FF.doWhat.value='3';chkForm2(document.FF,3,'<%=dbRS("patente")%>')" title="Activar"><img src="../images/ok.png" alt="Desactivar"></a>
			<%End IF%>	
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
			<td colspan ="3"> No hay vehiculos asoaciados</a></td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
