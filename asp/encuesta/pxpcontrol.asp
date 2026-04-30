<%@ Language=VBScript %>
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 

'Request.ServerVariables("REMOTE_HOST")
'Request.ServerVariables("REMOTE_ADRR")
'Request.ServerVariables("HTTP_HOST")
'Request.ServerVariables("HTTP_REFERER")
'response.write Request.ServerVariables("X_FORWARDED_FOR")
'response.write Request.ServerVariables("REMOTE_ADRR")
'response.write Request.ServerVariables("REMOTE_HOST")
'response.write Request.ServerVariables("X_FORWARDED_FOR_X86")
'response.write Request.ServerVariables("X_FORWARDED_FOR_X64")
'Response.Write(Request.ServerVariables("HTTP_X_FORWARDED_FOR_X86"))
'Response.Write(Request.ServerVariables("HTTP_X_FORWARDED_FOR_X64"))
 

Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor","exec","execute",_
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
 
if doWhat>="0" Then
	hdrid = clng("0"&objRequest("hdrid"))
	nrocliente = objRequest("nrocliente") 
	hdridfirst =  clng("0"&objRequest("hdridfirst") )
	nroclientefirst =  clng("0"&objRequest("nroclientefirst") )
	armadorid =  clng("0"&objRequest("armadorid") )
	controladorid =  clng("0"&objRequest("controladorid") )
	canterrores =  clng("0"&objRequest("canterrores") )
End If
 

'submit_logout = Request.Form("submit_logout")
If  Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If  

If doWhat = "-1"  Then 
	response.redirect "../menu.asp"
End If
 

Set objRequest = Nothing
%>
<HTML>
<HEAD>
<TITLE>ILLANES HNOS SRL</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css"> 
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


	<form name="FF" method="post" action="pxpcontrol.asp">
	
 
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	
    <div class="wrap" style="width: 70%" align="center">
	<div style="overflow-x: auto;">
		<h1>PEDIDO POR PEDIDO - CONTROL</h1>
<%
	if  doWhat="" then 
%>	 	 
 
	
		<div>
			<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
			<tr>
				<td>
				<span>Hoja de Ruta*</span>
				</td>
				
				<td>
					<input type="text" id="hdrid" name="hdrid" placeholder="60285" minlength="6" required value="<%=hdrid%>"><span class="labelError" id="hdriderrormsg"></span>
				</td>
			
			</tr>
			<tr>
				<td>
				<span>Nro Cliente*</span>
				</td>
				
				<td>
					<input type="text" id="nrocliente" name="nrocliente" placeholder="1245" minlength="7" required value="<%=nrocliente%>"><span class="labelError" id="nroclienteerrormsg"></span>
				</td>
				<td>
				</td>
			</tr>
			<tr>
				
				<td>
					<input type="button" class="btn1" value="Buscar" onClick="chkForm(this.form, 1)" style="width:80px;">
				</td>
				<td><input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;"></td>
			</tr>
			</table>
		</div>

<%
	End If 
	
	if doWhat="1" then 
	%>
	<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">
	<input type="hidden" name="hdridfirst" value="<%=hdrid%>">
	<input type="hidden" name="nroclientefirst" value="<%=nrocliente%>">
	<input type="hidden" name="hdrid" value="<%=hdrid%>">
	
	<!--#include file="./includes/pxpBuscar_HDR_Cliente.asp" -->

		
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,'')">

	<%
	end if 
	
	
	if doWhat="2" then 
		%> 


		<input type="hidden" name="nroclientefirst" id="nroclientefirst" value="<%=nroclientefirst%>">	 
		<input type="hidden" name="hdridfirst" id="hdridfirst" value="<%=hdridfirst%>">	 
		
		<input type="hidden" name="hdrid" value="<%=hdrid%>">
		<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">
		<span  class="pagetitle">HDR: <%=hdrid%></span>
		<span  class="pagetitle">Cliente: <%=nrocliente%></span>
		<input type="button" class="btn1" value="Volver" onClick="chkVolver(this.form, 1,<%=hdridfirst%>,<%=nroclientefirst%>)"> 

		<!--#include file="./includes/pxpArmadorControlador.asp" --> 
	 
		
		<div align="center">
			<br/> 
			
			<input type="button" class="btn1" value="Guardar" onClick="chkForm(this.form, 3)">		
		</div>
		 
		
		
		<% 
	End If	
								
	if doWhat>="3" Then
		%> 

		<!--#include file="./includes/pxpGuardarArmadorControlador.asp" --> 
		<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">	 
		<input type="hidden" name="hdrid" value="<%=hdrid%>">
		<input type="hidden" name="nroclientefirst" id="nroclientefirst" value="<%=nroclientefirst%>">	 
		<input type="hidden" name="hdridfirst" id="hdridfirst" value="<%=hdridfirst%>">	 
		<div align="center" class="textSuccess">Sus cambios fueron guardados</div>
		<div align="center">
			<br/> 
			<input type="button" class="btn1" value="Volver" onClick="chkVolver(this.form, 1,<%=hdridfirst%>,<%=nroclientefirst%>)">  
		</div>
		
		<%
 	End If				
%>					
			
    </div>
			
    </div>
	
	
	
	</Form>
	
<script language="javascript">
function chkControlar(Fm,prm,param1,param2){ 
	Fm.doWhat.value = prm;
	Fm.hdrid.value=param1;
	Fm.nrocliente.value=param2;
	Fm.submit(); 
}
function chkVolver(Fm,prm,param1,param2){ 
	Fm.doWhat.value = prm;
	Fm.hdrid.value=param1;
	Fm.nrocliente.value=param2;
	Fm.submit(); 
}

function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
} 

function chkForm(Fm,prm){
	var msg  = "";
	var question  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;


	if(prm == 1){  
		var input = document.getElementsByName("hdrid")[0].value;
		var cliente = document.getElementsByName("nrocliente")[0].value;
	 
		var pattern = /^[0-9]+$/;  
		if ((input) == "" && cliente== ""){
			msg = "Debe ingresar o HDR o Cliente";
			document.getElementById("hdriderrormsg").innerHTML = msg;
			document.getElementById("nroclienteerrormsg").innerHTML = msg;
		}
		else {
			
			if(!input.match(pattern) && input!="") {     
			  msg = "Debe ingresar una hoja de ruta válida";
			  document.getElementById("hdriderrormsg").innerHTML = msg;
			}
			if ((msg) == ""){
				var str = document.getElementById("nrocliente").value;
				if(  !str.match(pattern) && str!=""){
					msg = "Debe ingresar un cliente válido";
					document.getElementById("nroclienteerrormsg").innerHTML = msg;
				}
			}
		}
	}
	 	 
	if(prm == 3){
		var armador = document.getElementById("ArmadorID").value;
		var controlador = document.getElementById("ControladorID").value; 
		var canterrores = document.getElementById("canterrores").value; 
		
		if(armador==0){
			msg = "Debe seleccionar un armador";
			document.getElementById("armadoridmsg").innerHTML = msg;
		}
		if(canterrores>10){
			question = "Esta seguro que desea ingresar " + canterrores; 
			
			var answer = window.confirm(question); 
			if  (!(answer)) {
				return true;
			}
		} 
		if(controlador==0){
			msg = "Debe seleccionar un controlador";
			document.getElementById("controladoridmsg").innerHTML = msg;
		}  

 
	}
	  
	if(msg != "") alert(msg);
	else{ 
	 
		Fm.doWhat.value = prm;
		Fm.submit(); 
	}
}
 
  	 
function resetForm(Fm) {
	Fm.idFactura.value = "0";
	Fm.doWhat.value = "";
	Fm.submit(); 
}  
  	 
	 
 

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


