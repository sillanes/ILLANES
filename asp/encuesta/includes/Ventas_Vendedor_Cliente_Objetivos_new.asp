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


	<%If  NOT dbRS.EOF  Then%>
		<%DO UNTIL dbRS.EOF%>
	
	<div class="itemTD11">
		<span class="pagetitle"><%=dbRS("Nombre")%></span>

		<%If  dbRS("uploaded") = 1 AND dbRS("isValid") = 0  Then%>
		<img src="../images/reloj2.png" alt="En Proceso">
		<%elseif dbRS("uploaded") = 1 AND dbRS("isValid") = 1 Then %>
		<img src="../images/ok.png" alt="Validado">
		<%elseif dbRS("uploaded") = 1 AND dbRS("isValid") = 2 Then %>
		<img src="../images/del.png" alt="Descartado"> <a href="javascript:chkForm3(document.FF3,2,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Subir Foto"><img src="../images/upload.png" alt="Subir"></a>
		<%else%>
		<a href="javascript:chkForm3(document.FF3,2,'<%=ClienteID%>','<%=clientenombre%>','<%=dbRS("ObjetivoID")%>','<%=dbRS("Nombre")%>')" title="Subir Foto"><img src="../images/upload.png" alt="Subir"></a>
		<%End If%>


	</div>
		<%
	 
 			dbRS.MoveNext
		LOOP
		%>
	<%
	Else
	%>
		
		No tiene objetivos asociados
		
	<%
	End If
	
	%>
	
</table>

<span class="pagetitle"><img src="../images/ok.png" alt="Validado"> Foto validada por supervisor</span>
<span class="pagetitle"><img src="../images/upload.png" alt="Subir"> Subir una foto</span>
<span class="pagetitle"><img src="../images/del.png" alt="Subir"> Foto descartada por supervisor</span>
<span class="pagetitle"><img src="../images/reloj2.png" alt="Subir"> Foto Subida en espera de aprobacion</span>

</div>

<% 
Set dbRS = Nothing
%>
