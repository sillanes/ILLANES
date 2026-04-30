<%
  
	If retAction = 0 Then
%>
	<div><FONT FACE=”ARIAL”><img src="./images/errors.png" width="30" height="30"> Lo sentimos, no hemos encontrado sus datos<BR></FONT></div>


<%	
	End If
	If retAction = 1 Then
%>

	<div align="center" class="pagetitle">
		<span  align="center">Su reclamo <%=retReclamo%> ah sido actulizado</span>
		<span  align="center">Nos comunicaremos a la brevedad</span>
	</div>


<%	
	End If
	If retAction = 2 Then
%>

	<div align="center"  class="pagetitle">
		<span align="center">Su reclamo ah sido generado bajo el numero: <%=retReclamo%></span>
		<span align="center">Nos comunicaremos a la brevedad</span>
	</div>

<%	
	End If
   
%>
<div align="center" class="pagetitle">
<input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)">
</div>
<%  
Set dbRS = Nothing
%>
