
<%

tel=RDTelefono
cli=RDClienteID
rsoc=RDClienteNombre

%>
<div>
	<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
	<tr>
		<td>
		<span>Cliente:</span>
		</td>
		<td>
		<span class="labelSuccessfully"><%=rsoc%></span>
		</td>
	
	</tr>
	<tr>
		<td>
		<span>Telefono*</span>
		</td> 
		<td>
			<input type="text" id="telefonoid" name="telefonoid" placeholder="(299)536-0747" minlength="6" required value="<%=tel%>"><span class="labelError" id="telefonoiderrormsg"></span>
		</td>
		<td>
		</td>
	</tr>

	</table>
</div>