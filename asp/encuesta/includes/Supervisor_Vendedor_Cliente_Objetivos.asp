<%
'ON ERROR RESUME NEXT
totalPendinetes = 0 
sSQL = "EXEC super.[Vendedor_Clientes_Objetivo_Sel] " &"'"& VendedorID & "'," &  ClienteID & "," & PeriodoID 
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
		<td colspan="3" height="20">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tbody><tr class="tableHeader">
				<td align="left" colspan="3"><span class="formHeader">Objetivos</span></td>
				</tr>
			</tbody>
		</table>
		</td>
	</tr>
	</tbody>
	<tr align="center" class="columnTop">
		<td height="20" ><b>Objetivo</b></td>
		<td height="20" ><b>Vendedor</b></td>
		<td height="20" ><b>Supervisor</b></td>
	</tr>
	

	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
			
	<tr class="itemTD11">
		<td onclick="ShowComments('<%=dbRS("descripcion")%>')"><%=dbRS("Nombre")%><a href="javascript:return false;" title="<%=dbRS("descripcion")%>"><img src="../images/texto.png"></a> </td>
		
		<%If  dbRS("uploaded") = 1 AND dbRS("isValid") = 0  Then%>
		<td><img src="../images/reloj2.png" alt="En Proceso"> </td>
		<%elseif dbRS("uploaded") = 1 AND dbRS("isValid") = 1 Then %>
				<td><img src="../images/ok.png" alt="Validado"></td>
		<%elseif dbRS("uploaded") = 1 AND dbRS("isValid") = 2 Then %>
		<td onclick="ShowComments('<%=dbRS("ToolTip")%>')"><a href="javascript:return false;" title="<%=dbRS("ToolTip")%>"><img src="../images/del.png" alt="Descartado"></a> </td>
		<%else%>
		<td></td>		
		<%End If%>
		 
		<%If  dbRS("uploaded") = 1 Then %>
		<td>
		<a href="javascript:chkForm4(document.FF3,3,1,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Validar Foto"><img src="../images/control.png" alt="Validar"></a>
		<a href="javascript:chkForm4(document.FF3,3,2,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Descartar Foto"><img src="../images/basura.png" alt="Descartar"></a>	
		<a href="" onClick="return fun('<%=dbRS("NombreArchivo")%>');" title="Ver Foto"><img src="../images/eye.png" alt="Ver Factura" id="verfoto" ></a>		
		</td>
		<%else%>
		<td></td>		
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

<span class="pagetitle">Acciones Supervisor</span>
<span class="littleText"><img src="../images/control.png" alt="Validado">Validar Foto <img src="../images/basura.png" alt="Subir">Descartar foto <img src="../images/eye.png" alt="Subir"> Ver Foto</span> 
 
<br/><br/><br/>
<span class="pagetitle">Vista Vendedor</span>
<span class="littleText"><img src="../images/ok.png" alt="Validado">Foto validada <img src="../images/del.png" alt="Subir">Foto Descartada <img src="../images/reloj2.png" alt="Subir"> Foto pendiente de aprobación</span>


</div>

<% 
Set dbRS = Nothing
%>
