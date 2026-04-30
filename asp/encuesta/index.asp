<%@ Language=VBScript %>
<%
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_encuesta.asp" --><%
dbCon.CommandTimeout = 0 


Set objRequest = Request.Form
doWhat = objRequest("doWhat")
StartDate = Replace(objRequest("StartDate"),".","/")

idCliente = CLng("0" & objRequest("idCliente"))
idFactura = CLng("0" & objRequest("idFactura"))

' variales para buscar en arrys
Dim PreguntaID, PreguntaNombre, PreguntaAbierta
PreguntaID = 1 
PreguntaNombre = 2 
PreguntaAbierta = 3

Dim ClienteNombre 
ClienteNombre = ""
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


	<form name="FF" method="post" action="index.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="idCli" value="<%=idCliente%>">
	<input type="hidden" name="idFactura" value="<%=idFactura%>">
	
 
	
    <div class="wrap">
		 <!--#include file="./includes/getEncuestaEncabezado.asp" -->
<%
	if  doWhat="" then 
%>	 	
		<div>
			<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
			<tr>
			<td>
			<span>Cliente</span>
			</td>
			
			<td>
				<input type="text" id="CliNum" name="idCliente" placeholder="1245" minlength="3" required value="<%=idCliente%>"><span id="error"></span>
			</td>
			<td>
				<input type="button" class="btn1" value="Buscar" onClick="chkForm(this.form, 1)" style="width:80px;">
			</td>
			</tr>
			</table>
		</div>

 
<%
	End If

	if doWhat="1" then 
%>	 
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


	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%">
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
		
		 
<%

	
	%><!--#include file="./includes/getEncuestaPreguntas.asp" --><%
	%><!--#include file="./includes/getPreguntasRespuestas.asp" --><%


		IF VarType(aPreguntas) < 8192 Then
			response.write("<h2>") 
			response.write( "<span class='itemTD11'>No se encontraron datos para la información ingresada. Por favor comuniquse con atención de reclamos a <a href=""mailto:reclamos@illanes.com.ar"">reclamos@illanes.com.ar</a><span>") 
			response.write("</h2>") 
		Else
		
			jMax = ubound(aPreguntas, 2)
			for j = 0 to jMax
				response.write("<h2>") 
				response.write( aPreguntas(PreguntaID,j) & "") 
				
				response.write("</h2>") 
				
			next

		End If

	end if 	
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
		Fm.submit(); 
	}
}


function chkForm(Fm,prm){
	var msg  = "";
	if(prm == 1){
		var str = document.getElementById("CliNum").value;
		if(str.length<4)
			msg = "El nÃºmero de cliente debe ser mayor a cuatro dÃ­gitos";
		
		if ((msg) == ""){
			if(!numberValidationFunc())
				msg = "Ingrese un cÃ³digo de cliente vÃ¡lido";
		}
			
	}
	 
	if(prm == 2){  
		Fm.StartDate.value;
	}	 

	if(prm == 3){  
			
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

  	 
</script>

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%>



