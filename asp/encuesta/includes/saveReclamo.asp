<%

' @FechaFact INT,
' @ClienteID INT,
' @FacturaID INT,
' @Respuesta1 VARCHAR(2), -- pedido cerrado
' @Respuesta2 VARCHAR(2), -- falto producto
' @Respuesta3 VARCHAR(2), -- produto mal estado
' @text1 VARCHAR(500), -- texto falto producto
' @text2 VARCHAR(500), -- texto producto mal estado
' @comentarios VARCHAR(500),
' @email VARCHAR(100),
' @telefono VARCHAR(20) = NULL

Dim retReclamo, retAction

sSQL = "EXEC wsp_Reclamo_upd "& CLng("0"&dbStartDate) & ","  & CLng("0"&idCliente) & "," & CLng("0"&idFactura) & ",'" & pedcerrado  & "','" & pedfalta & "','" & pedmal  & "','" & txtpedfalta  & "','" & txtpedmal & "','" & pedcomentarios  & "','" & email & "','" & telefono & "'"
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 or dbRS.EOF  THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

If  NOT dbRS.EOF  Then
	retReclamo = dbRS("ReclamoGenerado")
	retAction = dbRS("Accion") 
	
	If retAction = 0 Then
%>
	<div><FONT FACE=”ARIAL”><img src="./images/errors.png" width="30" height="30"> Lo sentimos, no hemos encontrado sus datos<BR></FONT></div>


<%	
	End If
	If retAction = 1 Then
%>

	<div align="center" class="pagetitle">
		<span  align="center">Su reclamo <%=retReclamo%> ha sido actulizado.</span><br>
		<span  align="center">Nos comunicaremos a la brevedad</span>
	</div>


<%	
	End If
	If retAction = 2 Then
%>

	<div align="center"  class="pagetitle">
		<span align="center">Su reclamo ha sido generado bajo el numero: <%=retReclamo%>.</span><br>
		<span align="center">Nos comunicaremos a la brevedad</span>
	</div>

<%	
	End If
	
	If retAction = 3 Then
%>

	<div align="center"  class="pagetitle">
		<span align="center">Ya pose un reclamo en proceso con el número <%=retReclamo%>.</span><br>
		<span align="center">Nos comunicaremos a la brevedad</span>
	</div>

<%	
	End If	
	
	If retAction = 4 Then
%>

	<div align="center"  class="pagetitle">
		<span align="center"><%=retReclamo%></span><br> 
	</div>

<%	
	End If		
	
	if retAction = 99 Then
		RESPONSE.CLEAR
		response.redirect "./error.asp"
	End If
   
%>
<div align="center" class="pagetitle">
<input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)">
</div>
<% 
End IF
Set dbRS = Nothing
%>
