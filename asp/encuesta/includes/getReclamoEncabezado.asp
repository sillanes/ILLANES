<%

sSQL = "EXEC wsp_Reclamos_Sel 1" 
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False

IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If


   
%>


<%If  NOT dbRS.EOF  Then%>
<div>
	<h1><%=dbRS("name")%></h1>
	<%If  (dbRS("showdetails") = 1) Then%>
	<h2><%=dbRS("details")%></h2>
	<%End if%>
</div>

<%
Else
%>
	<div class="itemTD11">
		<h2> No se encuentrar registros, verifique seleccion o comuniquese con nuestro centro de atención de reclamos a  <a href="mailto:email@example.com">reclamos@illanes.com.ar</a> </h2>
	</div>
<%
End If
 
Set dbRS = Nothing
%>
