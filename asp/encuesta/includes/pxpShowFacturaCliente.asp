	<%
	sSQL = "EXEC pxp.usp_Buscar_Armador_Controlador " & CLng("0" & facturaid2) & "," & CLng("0"&nrocliente2)
	'Response.write(sSQL)
	Set dbRS=Server.CreateObject("ADODB.Recordset")
	dbRS.open sSQL, dbCon
	IsParentInactive = False
	IF ERR.NUMBER <> 0 THEN
		RESPONSE.CLEAR
		response.redirect "./error.asp"
	End If

	%>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%> 
		
		<h1>DATOS SELECCIONADOS</h1>
		<table>
		<div class="pagetitle">Factura: <%=dbRS("FacturaID")%></div>
		<div class="pagetitle">Cliente: <%=dbRS("ClienteNombre")%></div>
		<div class="pagetitle">Razon Social: <%=dbRS("RazonSocial")%></div>
		<div class="pagetitle">Armador: <%=dbRS("Armador")%></div>
		<div class="pagetitle">Controlador: <%=dbRS("Controlador")%></div> 
		
		</table>
		<input type="hidden" name="FechaFact" value="<%=dbRS("FechaFact")%>">
	
		<%
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<h1>SIN DATOS</h1>
		
	<%
	End If
	
	%>	
		
	<br/>  
 
		
	


<% 
Set dbRS = Nothing
%>
