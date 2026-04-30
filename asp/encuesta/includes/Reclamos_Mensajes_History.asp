<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Reclamos_Mensajes_Sel '"& idReclamo & "'"
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
		<td colspan="2"height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="2" align="left"><span class="formHeader">MENSAJES ENVIADOS</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"><b>Fila<br/></td>
		<td height="20" width="95%"><b>Mensaje<br/></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("MensajeItem")%></td>
		<td><textarea readonly id="textarea<%=dbRS("MensajeItem")%>" name="txtmsg<%=dbRS("MensajeItem")%>" style="overflow:auto;resize:none" wrap="hard" rows="4" cols="120" maxlength="500"><%=dbRS("Mensaje")%></textarea></td>
	</tr>
	
		<%
			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="2"> No hay mensajes anteriores</a>.</td>
		</tr>
		
	<%
	End If
	
	%>
	
</table>

<% 
Set dbRS = Nothing
%>
