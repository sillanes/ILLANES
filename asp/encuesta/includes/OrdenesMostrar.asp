<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Ordenes_Sel "& OrdenID
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
		<td colspan="3"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="3" align="left"><span class="formHeader">ORDENES EMPRETIENDA</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody> 
	<tr align="center" class="columnTop" > 
		<td height="20" width="5%"><b>Fecha<br/></td> 
		<td height="20" width="5%"><b>Cliente<br/></td>
		<td height="20" width="5%"><b>Monto<br/></td> 
	</tr>
		
	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
		

		<tr class="itemTD11"> 
			<td><%=dbRS("Fecha")%></td> 
			<td><%=dbRS("Cliente")%></td> 
			<td>$<%=FormatNumber(dbRS("Monto"),2)%></td>   
		</tr>

		 
	<%
			'idMonto = replace(dbRS("Monto"),",",".") 
			Ok = dbRS("OK")
			NroRecibo = dbRS("NroRecibo")
			NroCliente = dbRS("NroCliente")
			dbRS.MoveNext
		LOOP
		
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="3"> No existe la ordenes</td>
		</tr>
		
	<%
	End If
	 
%>
</table>


<div>
	<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
	 
	<tr>
		<td>
		<span>Ok</span>
		</td>
		
		<td>
			<input type="text" id="inputok" name="inputok"   required value="<%=ok%>"> 
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td>
		<span>NroRecibo</span>
		</td>
		
		<td>
			<input type="text" id="NroRecibo" name="NroRecibo"   required value="<%=NroRecibo%>"> 
		</td>
		<td>
		</td>
	</tr>


	</table> 
		
<%
Set dbRS = Nothing
%>
