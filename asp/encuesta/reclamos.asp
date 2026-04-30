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
StartDate = Replace(objRequest("StartDate"),".","/")

idCliente = CLng("0" & objRequest("idCliente"))
idFactura = CLng("0" & objRequest("idFactura"))

Dim  RowID, pedcerrado,pedfalta,pedmal,pedcommentarios
Dim valpedcerrado,valpedfalta,valpedmal
Dim ClienteNombre 
ClienteNombre = ""


if doWhat>="0" Then
	email = objRequest("email")
	telefono = objRequest("telefono")
End If


if doWhat>="1" Then
	if idCliente  = 0 Then	
		idCliente = objRequest("idCli")
	End If
	%><!--#include file="./includes/getNombreCliente.asp" --><%
End If

if doWhat>="2" Then

	idCliente = objRequest("idCli")
	dbStartDate = year(StartDate)  & right("0"&month(StartDate),2) &  right("0"&day(StartDate),2)
End If

if doWhat="4" Then
	result = 0
	idCliente = objRequest("idCli")
	dbStartDate = year(StartDate)  & right("0"&month(StartDate),2) &  right("0"&day(StartDate),2)
	idFactura = CLng("0" & objRequest("idFactura"))
	
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


	<form name="FF" method="post" action="reclamos.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="idCli" value="<%=idCliente%>">
	<input type="hidden" name="idFactura" value="<%=idFactura%>">
	
 
	
    <div class="wrap">
		 <!--#include file="./includes/getReclamoEncabezado.asp" -->
<%
	if  doWhat="" then 
%>	 	
 
	 
		<div>
			<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
			<tr>
				<td>
				<span>Email*</span>
				</td>
				
				<td>
					<input type="text" id="email" name="email" placeholder="ejemplo@mail.com.ar" minlength="4" required value="<%=email%>"><span class="labelError" id="emailerrormsg"></span>
				</td>
			
			</tr>
			<tr>
				<td>
				<span>Telefono*</span>
				</td>
				
				<td>
					<input type="text" id="telefono" name="telefono" placeholder="1245" minlength="7" required value="<%=telefono%>"><span class="labelError" id="telefonoerrormsg"></span>
				</td>
				<td>
				</td>
			</tr>
			<tr>
				
				<td>
					<input type="button" class="btn1" value="Siguiente" onClick="chkForm(this.form, 0)" style="width:80px;">
				</td>
				<td></td>
			</tr>
			</table>
		</div>
		<br/>
		<span class="labelError"> Estimado cliente, los reclamos deben realizarse dentro de las 72hs hábiles desde la fecha de facturación</span>
		<br/>

<%
	End If
	if  doWhat="0" then 
%>	 	
 
		<input type="hidden" name="email" value="<%=email%>">
		<input type="hidden" name="telefono" value="<%=telefono%>">
	

		<div>
			<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">

			<tr>
				<td>
				<span>Nro Cuenta:</span>
				</td>
				<td>
					<input type="text" id="CliNum" name="idCliente" placeholder="1245" minlength="3" required value="<%=idCliente%>"><span class="labelError" id="error"></span>
				</td>
				<td>
					<input type="button" class="btn1" value="Buscar" onClick="chkForm(this.form, 1)" style="width:80px;">
				</td>
			</tr>

			<tr>
				<td colspan="2">
				<span class="smallText">* El numero de cliente/nro cuenta figura en su factura </span>
				</td>
				<td colspan="1">
				<a  href="" onClick="return fun();" title="Ver Factura"><img src="../images/eye.png" alt="Ver Factura" id="verfactura" ></a>
				
				</td>
			</tr>

			</table>
		</div>
		<br/>
		<span class="labelError"> Estimado cliente, los reclamos deben realizarse dentro de las 72hs hábiles desde la fecha de facturación</span>
		<br/>
		
 
<%
	End If

	if doWhat="1" then 
%>	 
		<input type="hidden" name="email" value="<%=email%>">
		<input type="hidden" name="telefono" value="<%=telefono%>">
		<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm table-width-lg">
 		  <tr>		
			<td valign="top" id="SDate_show"><SPAN title="Ingrese fecha de la Factura">Fecha</SPAN>:&nbsp;
			<input type="Text" name="StartDate" id="StartDate"maxlength="10" class="width-20" value="<%= StartDate%>" chktype="date|1">
			<img src="../images/calendar.gif" alt="Calendar" id="StartDate_img" >
			<input type="button" class="btn1" value="Buscar" onClick="chkForm(this.form, 2)" style="width:80px;">
			</td>
		  </tr>
		 </table>
 
	
<%	
	else
		%>
		<input type="hidden" name="StartDate" value="<%=StartDate%>">
		<%
	end if 
		
	if doWhat="2" then 
		%><!--#include file="./includes/getFacturasCliente.asp" --><%
	end if 
	
	
	if doWhat="3" then 
	%>

	<input type="hidden" name="email" value="<%=email%>">
	<input type="hidden" name="telefono" value="<%=telefono%>">

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
			<td align="left" width="60%"><%=ClienteNombre%></td>
			<td align="left" ><%=StartDate%></td>
			<td><%=idFactura%></td>
		</tr>
				 
					   
	</table>
	
	<div  class="pagetitle">
		<span align="left"> ¿Su pedido llegó cerrado?</span>
	 
		<label>
			<input type="radio" name="pedcerrado" value="Si" checked> Si
		</label>
		<label>
			<input type="radio" name="pedcerrado" value="No"> No
		</label> 
		<br/>
	</div>
		  
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
		<span>Dejanos tu comentario, tu opinion nos ayuda a mejorar</span>
		<textarea name="txtcomment" style="overflow:auto;resize:none" rows="8" cols="20" maxlength="500"></textarea>
	</div>

	<div align="center">
		<br/>
		<input type="button" class="btn1" value="Enviar Reclamo" onClick="chkForm(this.form, 4)">
		
	</div>
<%
  

	end if 	
	
					
	if doWhat>="4" Then
	%><!--#include file="./includes/saveReclamo.asp" --><%
 	End If				
%>					
				
    </div>
	
	
	
	</Form>
	
<script language="javascript">
function chkForm2(Fm,prm,param1){
	var msg  = "";
	if(prm == 3){ 
		Fm.idFactura.value=param1;
	}		
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.idFactura.value=param1;
		Fm.submit(); 
	}
}


function chkForm(Fm,prm){
	var msg  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;

	if(prm == 0){
		var address = document.getElementById("email").value;

	   if(!reg.test(address)) {     
		  msg = "Debe ingresar un email válido";
		  document.getElementById("emailerrormsg").innerHTML = msg;
	   }
	   if ((msg) == ""){
			var str = document.getElementById("telefono").value;
			if(str.length<7){
				msg = "Debe ingresar un telefono válido";
				document.getElementById("telefonoerrormsg").innerHTML = msg;
			}
		}
	}
	 
	if(prm == 1){
		var str = document.getElementById("CliNum").value;
		if(str.length<4)
			msg = "El número de cliente debe ser mayor a cuatro dígitos";
		
		if ((msg) == ""){
			if(!numberValidationFunc())
				msg = "Ingrese un código de cliente válido";
		}
			
	}
	 
	if(prm == 4){  
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

function numberValidationFunc()
{
 var input = document.getElementById("CliNum").value;
 var pattern = /^[0-9]+$/;
 if(input.match(pattern))
 {
   return true;
 }
 else
 {
  document.getElementById("error").innerHTML ="Ingrese un numero";
  return false;
 }
}

if ((document.FF.elements['doWhat'].value) == 1) {
	Calendar.setup({displayArea:"SDate_show", inputField:"StartDate", ifFormat:"dd/mm/y", button:"StartDate_img", singleClick:true});
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
  	 
function resetForm(Fm) {
	Fm.idFactura.value = "0";
	Fm.doWhat.value = "";
	Fm.submit(); 
}  
  	 
	 


function fun(){
	openWindow();
	return false;
}
function openWindow(){
if (document.getElementById) {
   w = screen.availWidth;
   h = screen.availHeight;
}  
var w = 480, h = 340;
var popW = 800, popH = 700;

var leftPos = (w-popW)/2;
var topPos = (h-popH)/2;



msgWindow = window.open('','popup','width=' + popW + ',height=' + popH + 
                         ',top=' + topPos + ',left=' + leftPos + ',       scrollbars=yes');

msgWindow.document.write 
    ('<HTML><HEAD><TITLE>FACTURA</TITLE></HEAD><BODY><FORM NAME="form1">' +
    '<img src="../images/factura.png"><br/>' +
    '<div align="center"><INPUT TYPE="button" VALUE="CERRAR"onClick="window.close();"></div></FORM></BODY>   </HTML>');
}

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


