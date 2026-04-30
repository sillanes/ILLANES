<%@ Language=VBScript %>
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_administracion.asp" --><%
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

TransferenciaID = objRequest("TransferenciaID") &""
idCliente = objRequest("idCliente") &""
cliente = objRequest("cliente") &""
monto = replace(objRequest("idMonto"),",",".")
rango = objRequest("rango") &""
rdrango = objRequest("rdrango") &""
txtStartRange = CLng("0" &objRequest("txtStartRange") )
txtEndRange =CLng("0" &objRequest("txtEndRange") )
startrd = CLng("0" &objRequest("startrd") )
endrd =CLng("0" &objRequest("endrd") )

if cliente<>"" Then
	idCliente = cliente 
End If
if rdrango<>"" Then 
	rango=rdrango
End If
if startrd<>"" Then 
	txtStartRange=startrd
End If
if endrd<>"" Then 
	txtEndRange=endrd
End If

 

If monto="" Then	
	monto=0
End If

if rango="No" Then
	irango=0
else
	irango=1
End If

if cliente<>"" Then
	idCliente = cliente
	rango=rdrango
End If
	
response.write "doWhat: " & doWhat
response.write "<br> idCliente: " & idCliente
response.write "<br> idMonto: " & objRequest("idMonto") 
response.write "<br> rango: " & rango
response.write "<br> txtStartRange: " &txtStartRange
response.write "<br> txtEndRange: " & txtEndRange
response.write "<br> USER: " & Session("currentUser")

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


	<form name="FF" method="post" action="imputartransferencias.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="Cliente" value="<%=idCliente%>">
	<input type="hidden" name="monto" value="<%=idMonto%>">
	<input type="hidden" name="rdrango" value="<%=rango%>"> 
	<input type="hidden" name="startrd" value="<%=txtStartRange%>"> 
	<input type="hidden" name="endrd" value="<%=txtEndRangeRange%>"> 
	<input type="hidden" name="TransferenciaID" value="<%=TransferenciaID%>"> 
	
 
	
    <div class="wrap" style="width: 70%" align="center">
	<h1>Transferencias</h1>
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
					<input type="text" id="idCliente" name="idCliente" placeholder="NOMBRE CLIENTE" minlength="4"><span class="labelError" id="clientemsg"></span>
				</td>
			
			</tr>
			<tr>
				<td>
				<span>Monto</span>
				</td>
				
				<td>
					<input type="text" id="idMonto" name="idMonto" placeholder="1245" minlength="4" ><span class="labelError" id="facturamsg"></span>
				</td> 
			</tr>
			<tr>
				<td colspan="2">

					<div  class="pagetitle">
						<span>¿Desea buscar importe por un rango?</span>
						<label>
							<input type="radio" name="rango" value="Si" onclick="mostrarRango()"> Si
						</label>
						<label>
							<input type="radio" name="rango" value="No" checked onclick="ocultarRango()"> No
						</label> 

					</div>
				</td>
	
			</tr> 
			<tr>
				<td colspan="2">
				<div id="iddesde" style="overflow:auto;resize:none;display:none"><span>Desde</span> <input type="text" name="txtStartRange" id="txtStartRange" value="<%=txtStartRange%>"> 
				</div>
				</td>
			</tr>
			<tr>
				<td colspan="2">
				<div id="idhasta" style="overflow:auto;resize:none;display:none"><span>Hasta</span><input type="text" name="txtEndRange" id="txtEndRange" value="<%=txtEndRange%>">
				</div>
				</td>
			</tr> 
			<tr>
				
				<td colspan="2" align="center">
					<input type="button" class="btn1" value="buscar" onClick="chkForm(this.form, 1)" style="width:80px;">
					<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
				</td>
				
			</tr>
			</table>
		</div>

<%
	End If
	if doWhat="1"then 
	%>
	
	
	<!--#include file="./includes/Transferencias_Buscar.asp" --> 
	 
 
 	<input type="button" class="btn1" value="Volver" onClick="resetForm(this.form, '')" style="width:80px;">
	
	 
<%
	End If
	If doWhat="2" Then
%>
		<!--#include file="./includes/TransferenciasMostrar.asp" --> 
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,1)" style="width:80px;">
		<input type="button" class="btn1" value="Guardar" onClick="chkForm4(this.form,3)" style="width:80px;"> 
		 

<%			
	End If
 
	If doWhat="3" Then
%>
		<!--#include file="./includes/TransferenciasMostrar.asp" --> 
	 
<%			
	End If
	
	 
%>

    </div>
	
	
	</Form>
	
<script language="javascript">
function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
}


function chkForm(Fm,prm){
	var msg  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;
  
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
function chkForm2(Fm,prm,param1){
	var msg  = "";
 	
	if(msg != "") alert(msg);
	else{ 
		Fm.doWhat.value = prm; 
		Fm.TransferenciaID.value=param1; 
		Fm.submit(); 
	}
}

function ocultarRango() {
  document.getElementById("idhasta").style.display = 'none'; 
  document.getElementById("iddesde").style.display = 'none'; 
}  

function mostrarRango() {
  document.getElementById("idhasta").style.display = 'block'; 
  document.getElementById("iddesde").style.display = 'block'; 
}  


</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


