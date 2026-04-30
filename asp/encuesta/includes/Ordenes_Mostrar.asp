<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Ordenes_view "& FileID
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
		<td colspan="25" height="2">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="25" align="left"><span class="formHeader">FileID: <%=FileID%></span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" width="5%"></td> 
		<td height="20" width="5%" colspan="8" style="border-style: solid;border-color: #96D4D4;"><b>Informacion Archivo</b></td> 
		<td height="20" width="5%" colspan="5" style="border-style: solid;;border-color: #FF0000;"><b>Datos Verificados</b></td>  
		<td height="20" width="5%" rowspan="2"><b>Observaciones</b></td>    
		<td height="20" width="5%" colspan="6"><b>Productos</b></td>    
		<td height="20" width="5%" colspan="4" style="border-style: solid;;border-color: #008000;"><b>Estados</b></td>   
	</tr> 	
	<tr align="center"  class="columnTop">
		<td height="20" width="5%"><b>#<br/></b></td>  
		
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Fecha</td>   
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Cliente</td>   
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Telefono</td>   
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Email</td>    
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Razon Social</td>   
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Tipo Documento</td>   
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">Numero Documento</td>   
		<td height="20" width="5%" style="border-style: solid;border-color: #96D4D4;">CUIT Archivo</td>  
		
		<td height="20" width="5%" style="border-style: solid;border-color: #FF0000;">CUIT valido</td>
		<td height="20" width="5%" style="border-style: solid;border-color: #FF0000;">Cliente America</td> 
		<td height="20" width="5%" style="border-style: solid;border-color: #FF0000;">Vendedor America</td>  
		<td height="20" width="5%" style="border-style: solid;border-color: #FF0000;">Factura</td>  
		<td height="20" width="5%" style="border-style: solid;border-color: #FF0000;">Razon Social<br/>Validado</td>   
		  
		<td height="20" width="5%";><b>Bronce</b></td>  
		<td height="20" width="5%";><b>Plata</b></td>  
		<td height="20" width="5%";><b>Oro</b></td>   
		<td height="20" width="5%";><b>Platino</b></td>    
		<td height="20" width="5%";><b>Especial</b></td>   
		<td height="20" width="5%";><b>Estrella</b></td>   
		
		
		<td height="20" width="5%" style="border-style: solid;;border-color: #008000;"><b>Verificado?</b></td> 
		<td height="20" width="5%" style="border-style: solid;;border-color: #008000;"><b>Cliente</b></td>   
		<td height="20" width="5%" style="border-style: solid;;border-color: #008000;"><b>Vendedor</b></td> 
		<td height="20" width="5%" style="border-style: solid;;border-color: #008000;"><b>Factura</b></td> 
	</tr> 
	<tr>
	<td colspan="16"></td>
	</tr> 
	 
	<%If  NOT dbRS.EOF   Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("OrdenID")%></td>
		<td><%=dbRS("Fecha")%></td> 
		<td><%=dbRS("Cliente")%></td> 
		<td><%=dbRS("Telefono")%></td>
		<td><%=dbRS("Email")%></td> 
		<td><%=dbRS("Fact_Razonsocial")%></td>
		<td><%=dbRS("Fact_TipoDeDocumento")%></td>
		<td><%=dbRS("Fact_NumeroDeDocumento")%></td> 
		<td><%=dbRS("CUIT_Archivo")%></td> 
		<td><%=dbRS("CUIT_Valido")%></td>  
		<td><%=dbRS("ClienteID")%></td>  
		<td><%=dbRS("VendedorID")%></td> 
		<td><%=dbRS("FacturaNro")%></td>
		<td><%=dbRS("RazonSocial")%></td>
		
		<td>
<%
  Dim obs, shortObs, obsAttr
  obs = dbRS("Observaciones")

  ' Armamos preview de 30 caracteres
  If Len(obs) > 30 Then
      shortObs = Left(obs,30) & "..."
  Else
      shortObs = obs
  End If

  ' Escapamos para atributo HTML
  obsAttr = Replace(obs, """", "&quot;")   ' escapamos comillas dobles
  obsAttr = Replace(obsAttr, "'", "&#39;") ' escapamos comillas simples
%>
  <span style="cursor:pointer;color:blue;text-decoration:underline;"
        data-obs="<%=Server.HTMLEncode(obsAttr)%>"
        onclick="mostrarObs(this.getAttribute('data-obs'))">
    <%=Server.HTMLEncode(shortObs)%>
  </span>
</td> 

<td><%=dbRS("CajasBronce")%></td>
<td><%=dbRS("CajasPlata")%></td>
<td><%=dbRS("CajasOro")%></td>
<td><%=dbRS("CajasPlatino")%></td>
<td><%=dbRS("CajasEspecial")%></td>
<td><%=dbRS("CajasEstrella")%></td>
 
		<td>
		<%If dbRS("FacturaValidada")=-1  THEN%>
			
			<a href="javascript:return false;" title="Aguarde a que se verifiquen los datos"><img src="../images/reloj2.png" alt="Aguarde a que se verifiquen los datos"></a>
	
		<%else%>	

			<%If dbRS("FacturaValidada")=0  THEN%>
			
				<a href="javascript:return false;" title="Problema con cantidades"><img src="../images/warning.png" alt="Problema con cantidades"></a>
		
			<% Else %>
			
				<a href="javascript:return false;" title="Factura Valida"><img src="../images/Ok.png" alt="Factura Valida"></a>
		
		
			<%End If%>
		
	
		<%End IF%>	
		</td>	
		
		<td>
		<%If  dbRS("ClienteID") <> "" THEN%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:return false;" title="Completo"><img src="../images/ok.png" alt="Completo"></a>
		<%else%>
			&nbsp;&nbsp;&nbsp;
 			<a href="javascript:return false;" title="Faltan Datos"><img src="../images/warning.png" alt="Faltan Datos"></a>
		<%End IF%>		
		</td>
			
		<td>
		<%If dbRS("VendedorID")<>"" THEN%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:return false;" title="Completo"><img src="../images/ok.png" alt="Completo"></a>
		<%else%>
			&nbsp;&nbsp;&nbsp;
 			<a href="javascript:return false;" title="Faltan Datos"><img src="../images/warning.png" alt="Faltan Datos"></a>
		<%End IF%>		
		</td>		
		
		<td>
		<%If dbRS("FacturaNro")<>"" AND dbRS("FacturaArchivo")<>"" THEN%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:return false;" title="Completo"><img src="../images/ok.png" alt="Completo"></a>
			
		<%else%>
			&nbsp;&nbsp;&nbsp;
 			<a href="javascript:return false;" title="Faltan Datos"><img src="../images/warning.png" alt="Faltan Datos"></a>
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
			<td colspan ="26"> No hay ordenes para visualizar</td>
		</tr>
		
	<%
	End If
	%>
</table>

<% 
Set dbRS = Nothing
%>

<!-- Modal Observaciones -->
<div id="modalObs" style="display:none;position:fixed;top:50%;left:50%;
    transform:translate(-50%,-50%);background:#fff;padding:20px;
    border-radius:8px;box-shadow:0 0 10px rgba(0,0,0,0.5);z-index:1000;width:400px;max-height:300px;overflow:auto;">
  <h3>Observaciones</h3>
  <pre id="obsTexto" style="white-space:pre-wrap;"></pre>
  <button type="button" onclick="document.getElementById('modalObs').style.display='none'">
  Cerrar
</button>

</div>

<script>
function mostrarObs(texto) {
    document.getElementById("obsTexto").innerText = texto;
    document.getElementById("modalObs").style.display = "block";
}
</script>

