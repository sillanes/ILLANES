<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC ven.[Vendedor_Clientes_Objetivo_Sel] " &"'"& VendedorID & "'," &  ClienteID & "," & PeriodoID 
'Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

   
%>

<div style="overflow-x: auto;">
<table id="MyTable" border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable">
	<tbody> 
	<tr>
		<td colspan="2" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="2"><span class="formHeader">Objetivos</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" ><b>Objetivo</b></td>
		<td height="20" ><b>Accion</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td onclick="ShowComments('<%=dbRS("descripcion")%>')"><%=dbRS("Nombre")%><a href="javascript:return false;" title="<%=dbRS("descripcion")%>"><img src="../images/texto.png"></a> </td>
		<%If  dbRS("uploaded") = 1 AND dbRS("isValid") = 0  Then%>
		<td><img src="../images/reloj2.png" alt="En Proceso"><a href="javascript:chkForm3(document.FF3,2,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Subir Foto"><img src="../images/upload.png" alt="Subir"></a></td>
		<%elseif dbRS("uploaded") = 1 AND dbRS("isValid") = 1 Then %>
				<td><img src="../images/ok.png" alt="Validado"></td>
		<%elseif dbRS("uploaded") = 1 AND dbRS("isValid") = 2 Then %>
		<td ><a onclick="ShowComments('<%=dbRS("ToolTip")%>')" href="javascript:return false;" title="<%=dbRS("ToolTip")%>"><img src="../images/del.png" alt="Descartado"></a> <a href="javascript:chkForm3(document.FF3,2,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Subir Foto"><img src="../images/upload.png" alt="Subir"></a></td>
		<%else%>
		<td><a href="javascript:chkForm3(document.FF3,2,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Subir Foto"><img src="../images/upload.png" alt="Subir"></a></td>
		<%End If%>
		 
	</tr>
	
		<%
	 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		<tr class="itemTD11" >
			<td colspan ="2"> No tiene objetivos asociados</a>.</td>
		</tr> 
		
	<%
	End If
	
	%>
	
</table>

<span class="pagetitle"><img src="../images/ok.png" alt="Validado"> Foto validada por supervisor</span>
<span class="pagetitle"><img src="../images/upload.png" alt="Subir"> Subir una foto</span>
<span class="pagetitle"><img src="../images/del.png" alt="Subir"> Foto descartada por supervisor</span>
<span class="pagetitle"><img src="../images/reloj2.png" alt="Subir"> Foto subida en espera de aprobación</span>

</div>

<% 
Set dbRS = Nothing
%>
