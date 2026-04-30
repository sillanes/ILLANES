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
	facturaid = clng("0"&objRequest("facturaid"))
	nrocliente = objRequest("nrocliente") 
	facturaid2 = clng("0"&objRequest("facturaid2"))
	nrocliente2 = objRequest("nrocliente2")  
	armadorid =  clng("0"&objRequest("armadorid") )
	controladorid =  clng("0"&objRequest("controladorid") )
	canterrores =  clng("0"&objRequest("canterrores") )
	ResolucionMotivoID= clng("0"&replace(objRequest("ResolucionMotivoID"),"'",""))
End If

'response.write "ResolucionMotivoID-->" & ResolucionMotivoID
	
if doWhat="3" Then
	result = 0 
	'ResolucionMotivoID= replace(objRequest("ResolucionMotivoID"),"'","")
	FechaFact = objRequest("FechaFact") 
	
	' Empezando a obtener las preguntas
	RowID = 0
	pedcerrado = objRequest("pedcerrado")
	pedfalta   = objRequest("pedfalta")
	pedmal     = objRequest("pedmal") 
	pedcomentarios = objRequest("txtcomment")
	txtpedfalta = objRequest("txtpedfalta")
	txtpedmal = objRequest("txtpedmal")
	
	'response.write "pedcerrado-->" & pedcerrado
	'response.write "pedfalta-->" & pedfalta
	'response.write "pedmal-->" & pedmal
	'response.write "pedcomentarios-->" & pedcomentarios
	
	'for RowID = 0 to pedcerrado.length - 1
	'	if pedcerrado(RowID).checked then
	'		valpedcerrado = pedcerrado(RowID).value
	'	end if
	'next
	'
	'for RowID = 0 to document.FF.pedfalta.length - 1
	'	if  pedfalta(RowID).checked then
	'		valpedfalta =  pedfalta(RowID).value
	'	end if
	'next
	'
	'for RowID = 0 to document.FF.pedmal.length - 1
	'	if  pedmal(RowID).checked then
	'		valpedmal =  pedmal(RowID).value
	'	end if
	'next	
	
	
	
	
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


	<form name="FF" method="post" action="pxpcargarreclamo.asp">
	
 
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	
    <div class="wrap" style="width: 70%" align="center">
		<h1>CARGAR RECLAMO</h1>
<%
	if  doWhat="" then 
%>	 	 
 
	
		<div>
			<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
			<tr>
				<td>
				<span>Factura*</span>
				</td>
				
				<td>
					<input type="text" id="facturaid" name="facturaid" placeholder="60285" minlength="6" required value="<%=facturaid%>"><span class="labelError" id="facturaiderrormsg"></span>
				</td>
			
			</tr>
			<tr>
				<td>
				<span>Cliente*</span>
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
	
	if doWhat>="1" and doWhat<="2"  then 
	%>
	<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>"> 
	<input type="hidden" name="nrocliente2" value="<%=nrocliente2%>">
	<input type="hidden" name="facturaid2" value="<%=facturaid2%>">
	<input type="hidden" name="facturaid" value="<%=facturaid%>">
	
	<!--#include file="./includes/pxpGetFacturaCliente.asp" -->

	<%
	if doWhat ="1"  then 
	%>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,'')"> 
	<%End If%>


<%
	end if 
	
	if doWhat="2"  then  
		
%> 
	
	<!--#include file="./includes/pxpShowFacturaCliente.asp" -->
	
	<!--#include file="./includes/Reclamos_Resolucion_Motivos_FromCargarReclamo.asp" --> 		
	<br/>
	

	<span align="left"> ¿Su pedido llegó cerrado?</span>

	<label>
		<input type="radio" name="pedcerrado" value="Si" checked> Si
	</label>
	<label>
		<input type="radio" name="pedcerrado" value="No"> No
	</label> 
	<br/> 
	  
	<div  class="pagetitle">
	<span>¿Faltó algun producto?</span>

	<label>
		<input type="radio"  name="pedfalta" value="Si" onclick="mostrarPedFalto()"> Si
	</label>
	<label>
		<input type="radio" name="pedfalta" value="No" checked onclick="ocultarPedFalto()"> No
	</label> 

	<textarea name="txtpedfalta" id="txtpedfalta" style="overflow:auto;resize:none;display:none" rows="8" cols="20" maxlength="500"></textarea>
	<br/>
	</div> 

	<div  class="pagetitle">
	<span>¿Llego algo en mal estado?</span>
	<label>
		<input type="radio" name="pedmal" value="Si" onclick="mostrarPedMal()"> Si
	</label>
	<label>
		<input type="radio" name="pedmal" value="No" checked onclick="ocultarPedMal()"> No
	</label> 
	<textarea name="txtpedmal" id="txtpedmal" style="overflow:auto;resize:none;display:none" rows="8" cols="20" maxlength="500"></textarea>
	<br/>
	</div>

	<div  class="pagetitle">
	<span>Observaciones</span>
	<textarea name="txtcomment" style="overflow:auto;resize:none" rows="8" cols="20" maxlength="500"></textarea>
	</div>

	<div align="center">
	<br/>
	<input type="button" class="btn1" value="Volver" onClick="chkForm2(this.form, 1,<%=facturaid%>,<%=nrocliente%>)">
	<input type="button" class="btn1" value="Cargar Reclamo" onClick="chkForm(this.form, 3)">

	</div>	
	
<%
	end if 
	

					
	if doWhat>="3" Then
	%><!--#include file="./includes/pxpsaveReclamo.asp" --><%
 	End If		
%>
	
    </div>
	
	
	
	</Form>
	
<script language="javascript">
function chkControlar(Fm,prm,param1,param2){ 
	Fm.doWhat.value = prm;
	Fm.facturaid.value=param1;
	Fm.nrocliente.value=param2;
	Fm.submit(); 
}
function chkVolver(Fm,prm,param1,param2){ 
	Fm.doWhat.value = prm;
	Fm.facturaid2.value=param1;
	Fm.nrocliente2.value=param2;
	Fm.submit(); 
}

function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
} 

function chkForm2(Fm,prm,param1,param2){
	Fm.doWhat.value = prm;
	Fm.facturaid2.value=param1;
	Fm.nrocliente2.value=param2;
	Fm.submit();
} 

function chkForm(Fm,prm){
	var msg  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;


	if(prm == 1){  
		var input = document.getElementsByName("facturaid")[0].value;
		var cliente = document.getElementsByName("nrocliente")[0].value;
	 
		var pattern = /^[0-9]+$/;  
		if ((input) == "" && cliente== ""){
			msg = "Debe ingresar o Factura o Cliente";
			document.getElementById("facturaiderrormsg").innerHTML = "Debe ingresar una factura";
			document.getElementById("nroclienteerrormsg").innerHTML = "Debe ingresar un cliente";
		}
		else {
			
			if(!input.match(pattern) && input!="") {     
			  msg = "Debe ingresar una factura válida";
			  document.getElementById("facturaiderrormsg").innerHTML = msg;
			}
			if ((msg) == ""){
				var str = document.getElementById("nrocliente").value;
				if(  !str.match(pattern) && str!=""){
					msg = "Debe ingresar un cliente válido";
					document.getElementById("nroclienteerrormsg").innerHTML = msg;
				}
			}
		} 
		
		if ( (input == 0 && cliente== "") || (input == "" && cliente== 0) ) {
			if (cliente==""){
				msg= "Debe Ingresar un cliente ";
				document.getElementById("nroclienteerrormsg").innerHTML = msg;
			}else{
				msg= "Debe Ingresar una Factura ";
				document.getElementById("facturaiderrormsg").innerHTML = msg;
			}
		}
		
		if ( (input == 0 ) && (cliente== 0) ) {
			if (cliente==0){
				msg= "Debe Ingresar un cliente ";
				document.getElementById("nroclienteerrormsg").innerHTML = msg;
			}else{
				msg= "Debe Ingresar una Factura ";
				document.getElementById("facturaiderrormsg").innerHTML = msg;
			}
		}		
		
		
	}
	 	 
	if(prm == 3){  
		let elementoActivo = document.querySelector('input[name="pedfalta"]:checked');
		if(elementoActivo) {
			if (elementoActivo.value == "Si"){
				txt = document.getElementById("txtpedfalta").value;
				if (txt.length==0){
					msg = "Falta el comentario de la pregunta, ¿Faltó algun producto?";
				}

			}
		} else {
			msg = "Falta el comentario de la pregunta";
		}
		
		if (msg=="") {
			let elementoActivo = document.querySelector('input[name="pedmal"]:checked');
			if(elementoActivo) {
				if (elementoActivo.value == "Si"){
					txt = document.getElementById("txtpedmal").value;
					if (txt.length==0){
						msg = "Falta el comentario de la pregunta, ¿Llego algo en mal estado?";
					}

				}
			} else {
				msg = "Falta el comentario de la pregunta";
			}
		}
		
			
	}	 
	  
	if(msg != "") alert(msg);
	else{ 
	 
		Fm.doWhat.value = prm;
		Fm.submit(); 
	}
}
 
  	 
function resetForm(Fm) {
	Fm.doWhat.value = "";
	Fm.submit(); 
}  
  	 


function mostrarPedFalto() {
  document.getElementById("txtpedfalta").style.display = 'block'; 
}  


function ocultarPedFalto() {
  document.getElementById("txtpedfalta").style.display = 'none'; 
}  

function mostrarPedMal() {
  document.getElementById("txtpedmal").style.display = 'block'; 
}  


function ocultarPedMal() {
  document.getElementById("txtpedmal").style.display = 'none'; 
}  	 
 

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


