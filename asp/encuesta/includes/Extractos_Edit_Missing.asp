
<%
If CUIT_Archivo<>"" Then
	cuit=CUIT_Archivo
End If
%>
<div>
	<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
	<tr>
		<td>
		<span>CUIT*</span>
		</td>
		
		<td>
			<input type="text" id="cuitid" name="cuitid" placeholder="99-999999-99" minlength="6" required value="<%=cuit%>"><span class="labelError" id="cuitiderrormsg"></span>
		</td>
	
	</tr>
	<tr>
		<td>
		<span>Razon Social*</span>
		</td>
		
		<td>
			<input type="text" id="razonsocial" name="razonsocialnombre" placeholder="PEPE SRL" minlength="7" required value="<%=razonsocial%>"><span class="labelError" id="razonsocialerrormsg"></span>
		</td>
		<td>
		</td>
	</tr>
	<tr>
		<td>
		<span><b>Sin cuit asociado</b></span>
		</td>
		
		<td>
			<input type="checkbox" id="chksincuit"  name="chksincuit">
		</td>
		<td>
		</td>
	</tr>

	</table>
</div>