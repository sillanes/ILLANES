<%
'ON ERROR RESUME NEXT 
sSQL = "EXEC dbo.usp_Ordenes_view_missing " & FileID
' Response.write(sSQL)
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
		<td colspan="24" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td colspan="24" align="left"><span class="formHeader">FileID: <%=FileID%></span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td></td> 
		<td colspan="9" style="border:solid 1px #96D4D4;"><b>Informacion Archivo</b></td> 
		<td colspan="5" style="border:solid 1px #FF0000;"><b>Datos Verificados</b></td>
		<td colspan="6" style="border:solid 1px;"><b>Productos</b></td>  
		<td rowspan="2" style="border:solid 1px #008000;"><b>Accion</b></td>   
	</tr> 	
	<tr align="center"  class="columnTop">
		<td><b>#</b></td>  
		<td style="border:solid 1px #96D4D4;">Fecha</td>   
		<td style="border:solid 1px #96D4D4;">Cliente</td>   
		<td style="border:solid 1px #96D4D4;">Telefono</td>   
		<td style="border:solid 1px #96D4D4;">Email</td>     
		<td style="border:solid 1px #96D4D4;">Razon Social</td>   
		<td style="border:solid 1px #96D4D4;">Tipo Documento</td>   
		<td style="border:solid 1px #96D4D4;">Numero Documento</td>   
		<td style="border:solid 1px #96D4D4;">CUIT Archivo</td>   
		<td style="border:solid 1px #96D4D4;">Observaciones</td>  
		<td style="border:solid 1px #FF0000;">CUIT valido</td>
		<td style="border:solid 1px #FF0000;">Cliente America</td> 
		<td style="border:solid 1px #FF0000;">Vendedor America</td>  
		<td style="border:solid 1px #FF0000;">Razon Social Validado</td>
		<td style="border:solid 1px #FF0000;">Factura</td>
		<td style="border:solid 1px;">Bronce</td>
		<td style="border:solid 1px;">Plata</td>    
		<td style="border:solid 1px;">Oro</td>
		<td style="border:solid 1px;">Platino</td>    
		<td style="border:solid 1px;">Epecial</td>    
		<td style="border:solid 1px;">Estrella</td>    
	</tr> 
	<tr><td colspan="16"></td></tr> 
	 
	<%If  NOT dbRS.EOF   Then
		DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td width="5"><%=dbRS("OrdenID")%>
		    <input type="hidden" name="OrdenID_<%=dbRS("OrdenID")%>" value="<%=dbRS("OrdenID")%>">
		</td>
		<td><%=dbRS("Fecha")%></td> 
		<td><%=dbRS("Cliente")%></td> 
		<td><%=dbRS("Telefono")%></td>
		<td><%=dbRS("Email")%></td> 
		<td><%=dbRS("Fact_Razonsocial")%></td>
		<td><%=dbRS("Fact_TipoDeDocumento")%></td>
		<td><%=dbRS("Fact_NumeroDeDocumento")%></td> 
		<td><%=dbRS("CUIT_Archivo")%></td> 

		<td>
<% 
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
		<td><%=dbRS("CUIT_Valido")%></td>  
		<td><%=dbRS("ClienteID")%></td>  
		<td><%=dbRS("VendedorID")%></td> 
		<td><%=dbRS("RazonSocial")%></td>
		<td><%=dbRS("FacturaNro")%></td>
		<td><%=dbRS("CajasBronce")%></td>
		<td><%=dbRS("CajasPlata")%></td>
		<td><%=dbRS("CajasOro")%></td>
		<td><%=dbRS("CajasPlatino")%></td>
		<td><%=dbRS("CajasEspecial")%></td>
		<td><%=dbRS("CajasEstrella")%></td>
		
		<td>
 		<a href="javascript:document.FF.doWhat.value='5';chkForm3(document.FF,5,<%=FileID%>,<%=dbRS("OrdenID")%>)"  title="Modificar Faltantes"><img src="../images/control.png" alt="Modificar Faltantes" id="<%=dbRS("OrdenID")%>"></a>
		<a href="javascript:document.FF.doWhat.value='11';chkForm3(document.FF,11,<%=FileID%>,<%=dbRS("OrdenID")%>)"  title="Eliminar Orden"><img src="../images/del.png" alt="Eliminar Orden" id="<%=dbRS("OrdenID")%>"></a>
		</td>
	</tr>
	
		<%
			dbRS.MoveNext
		LOOP
	Else
	%>
		<tr class="itemTD11">
			<td colspan="24"> No hay ordenes para visualizar</td>
		</tr>
		
	<% End If %>
</table>
 

<% Set dbRS = Nothing %>
  
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