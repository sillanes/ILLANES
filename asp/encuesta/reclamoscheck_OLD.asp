<%@ Language=VBScript %> 
<!-- #include file="uploadhelper.asp" -->
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 

Dim BlackList, ErrorPage
BlackList = Array("cursor","exec","execute",_
				  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
'Note: We can include following keyword to make a stronger scan but it will also 
'protect users to input these words even those are valid input
'  "!", "char", "alter", "begin", "cast", "create",  
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
 

For Each s in Request.Form
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		' Redirect to an error page
		Response.Redirect(ErrorPage)
	End If
Next
%><!--#include virtual="./includes/sql-check.asp"--><%

 
Set objRequest = Request.Form
doWhat = objRequest("doWhat")
idReclamo = objRequest("idReclamo") 
txtresol  = objRequest("txtresolucion") 
subirarchivo = objRequest("subirarchivo")

submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If

'response.write "doWhat: " & doWhat
'response.write "<br> idReclamo: " & idReclamo
'response.write "<br> Subirarchivo: " & objRequest("subirarchivo")

  

Set objRequest = Nothing
%>
<HTML>
<HEAD>
<TITLE>ILLANES HNOS SRL</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />
<script type="text/javascript" src="../includes/calendar_cool.js"></script>
<script type="text/javascript" src="../includes/copy.js"></script>
    <style>
        li {
            cursor: pointer;
			margin: 15px 0;
        }
    </style>
 
  
</HEAD>
<body>


	<form name="FF" method="post" action="reclamoscheck.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="idReclamo" value="<%=idReclamo%>"> 
	<input type="hidden" name="txtresol" value="<%=txtresol%>"> 
 
	
    <div class="wrap" style="width: 55%">
		 
<%
	if  doWhat="" then 
%>	 	
		<h1>RECLAMOS PENDIENTES</h1> 
 
		<!--#include file="./includes/getReclamosPendientes.asp" -->
		 
		<span class="textInfoBig">Total Pendientes: <%=totalPendinetes%> </span>
		<input type="submit" name="submit_logout" value="Salir" />
<%
	End If 
	
	if doWhat="1" then 
	%>
	<h1>Reclamo: <%=idReclamo%></h1>
	<!--#include file="./includes/getReclamoDatos.asp" -->
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody>
		<tr>
			<td colspan="3" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="3"><span class="formHeader">Datos seleccionados</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="left" class="columnTop">
			<td height="20" width="60"><b>Cliente</b></td>
			<td height="20" width="30"><b>Fecha</b></td>
			<td height="20" width="30"><b>Nro Factura</b></td>
		</tr>
 

		<tr  align="left" class="itemTD11">
			<td align="left" width="60%"><%=dbRS("Cliente")%></td>
			<td align="left" ><%=dbRS("FechaFactura")%></td>
			<td><%=dbRS("FacturaNumero")%></td>
		</tr>
				 
					   
	</table>
	
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
		<tr>
			<td colspan="2" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="1" width="50%"><span class="formHeader">Reclamo</span></td>
					<td align="left" colspan="1" width="50%"><span class="formHeader">Resolucion</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
	</tbody>
	<tr>
	<td width="50%">
	<span  class="pagetitle" >
		<span align="left"> ¿Su pedido llegó cerrado?</span>
		<span align="left" class="textFail"><%=dbRS("PedidoCerrado")%> </span>
 
	</span>
		  
	<span  class="pagetitle">
		<span>¿Faltó algun producto?</span>
		<span align="left" class="textFail"><%=dbRS("FaltoProducto")%> </span>
		<%if dbRS("FaltoProducto") = "Si" Then %>
		<div align="left">
			<span class="textFail"><%=trim(dbRS("FaltoProductoTexto"))%> </span>
		</div> 
		<%End If%> 
	</span> 
	
	<span  class="pagetitle">
		<span>¿Llegó algo en mal estado?</span>
		<span align="left" class="textFail"><%=dbRS("ProductoMalEstado")%> </span>
		<%if dbRS("ProductoMalEstado") = "Si" Then %>
		<div align="left">
			<span class="textFail"><%=trim(dbRS("ProductoMalEstadoTexto"))%> </span>
		</div> 
		<%End If%>	 
	</span>
 
	<span class="pagetitle">
		<span>Comentario del cliente</span> 
		<div align="left">
			<span class="textFail"><%=trim(dbRS("Comentario"))%> </span>
		</div> 
	</span>
	</td>
		
	<td>
		<span  class="pagetitle">
			<span>Comentario hacia el cliente</span> 
			<textarea name="txtresolucion" style="overflow:auto;resize:none" wrap="hard" rows="12" cols="20" maxlength="500">Estimado <%=dbRS("NombreCliente")%>:</textarea>

		</span>
	</td>
	</tr>
	<tr>
	<td colspan="2">
	<div align="center">
		<br/>
		<input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)">
		<input type="button" class="btn1" value="Cerrar Reclamo" onClick="chkForm(this.form, 2)">		
	</div>
	</td>
	</tr>
	</table>
	
	
<%
  

	end if 	
	
					
	if doWhat>="2" Then
	'response.write "texto" & txtresol
%>	
	<!--#include file="./includes/getReclamoDatosSave.asp" -->
	<!--#include file="./includes/getReclamoDatos.asp" -->
	<h1>Reclamo: <%=idReclamo%></h1>
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody>
		<tr>
			<td colspan="3" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="3"><span class="formHeader">Datos seleccionados</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="left" class="columnTop">
			<td height="20" width="60"><b>Cliente</b></td>
			<td height="20" width="30"><b>Fecha</b></td>
			<td height="20" width="30"><b>Nro Factura</b></td>
		</tr>
 

		<tr  align="left" class="itemTD11">
			<td align="left" width="60%"><%=dbRS("Cliente")%></td>
			<td align="left" ><%=dbRS("FechaFactura")%></td>
			<td><%=dbRS("FacturaNumero")%></td>
		</tr>
				 
					   
	</table>
	
	<div align="center" class="textSuccess">El reclamo <%=idReclamo%> ha sido cerrado <br/><input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)"></div>
	

<%
 	End If				
%>					
				
    </div>
	 
	
	</Form>
	
<script language="javascript">
 
function chkForm2(Fm,prm,param1){
	var msg  = "";

	if(prm == 1){ 
		Fm.idReclamo.value=param1;
	}		
	if(msg != "") alert(msg);
	else{ 
		Fm.doWhat.value = prm; 
		Fm.idReclamo.value=param1; 
		Fm.submit(); 
	}
}


function chkForm(Fm,prm){
	var msg  = "";
	var txt = ""; 
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.submit(); 
	}
} 
function resetForm(Fm) {
	Fm.idReclamo.value = "";
	Fm.doWhat.value = "";
	Fm.submit(); 
}  
  	 
	 
 

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


