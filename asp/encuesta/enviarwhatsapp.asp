<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_WhatsAppapi.asp" --><%
dbCon.CommandTimeout = 0 
 

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
HojaDeRutaID = objRequest("HojaDeRutaID") 
hojaderutarow = objRequest("hojaderutarow") 
excluir = objRequest("excluir") 
telefonoid = objRequest("telefonoid")  
StartDate = Replace(objRequest("StartDate"),".","/")
fileupload = objRequest("cmd")
sSQL = objRequest("sSQL") 
nombrearchivo = objRequest("nombrearchivo")   


'Response.Write("chksincuit: "&chksincuit& "<br>"&VbCrLf)

'response.write "Extracto>" &  extracto
'response.write "hojaderutarow>" &  hojaderutarow


if doWhat="" or doWhat<="3" Then
	Session("FileUploaded") = ""
End If

'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
 

	<form  action="../includes/app_excel_whatsapp.asp" method="post" name="FEXCEL"> 
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="HojaDeRutaID" value="<%=HojaDeRutaID%>"> 
	<input type="hidden" name="hojaderutarow" value="<%=hojaderutarow%>"> 
	<input type="hidden" name="excluir" value="<%=excluir%>">  
	<input type="hidden" name="razonsocial" value="<%=razonsocialnombre%>"> 
	
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
 
	</form>
	

	<form name="FF" method="post" action="enviarWhatsApp.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="HojaDeRutaID" value="<%=HojaDeRutaID%>"> 
	<input type="hidden" name="hojaderutarow" value="<%=hojaderutarow%>"> 
	<input type="hidden" name="excluir" value="<%=excluir%>"> 
	<input type="hidden" name="cuit" value="<%=cuitid%>"> 
	<input type="hidden" name="razonsocial" value="<%=razonsocialnombre%>"> 
	
	
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
	
 
	
    <div class="wrap" style="width: 95%" align="center">
	
	
 <%
	if doWhat="" or doWhat="0"then 
	%>
	
	<!--#include file="./includes/WhatsAPP_Estado_ServicioFTP.asp" --> 
	<%if serviceshttp=1 THEN %>
	<div align="right"><h2>ESTADO SERVICIO FTP <img src="../images/ok.png" alt="Servicio FTP Activo"></h2></div>
	<%else %>
	<div align="right"><h2>ESTADO SERVICIO FTP <img src="../images/excluir.png" alt="Servicio FTP Activo"></h2></div>
	<%End If %>
 	<h1>Factura WhatsApp</h1>
  
  
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_EnProceso.asp" --> 
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Procesados.asp" --> 

	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
 <%
	End If
%>	
	
<%
	if doWhat="1"then 
	'HDR_Mostrar
	%>
	
	<h1>Hoja de Ruta</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Mostrar.asp" --> 
	
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Mostrar_Excluidos.asp" --> 
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	

	
<%
	if doWhat="2"then 
	'HojaDeRuta_Enviar
	%>
 
	<h1>Hoja De Ruta - Enviar</h1>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<br/><br/>
	<!--#include file="./includes/WhatsApp_HojaDeRuta_Enviar.asp" --> 
	<span class="textSuccess">La hoja de ruta: <%=HojaDeRutaID%> fue ingresada para procesar</span><br/>
	<br />
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	
	
	
<%
	if doWhat="3"then 
	'WhatsApp_Eliminar
	%>
 
	<h1>Hoja De Ruta - Eliminar</h1>
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Eliminar.asp" --> 
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<br /><br />
	<span class="textSuccess">La hoja de ruta: <%=HojaDeRutaID%> fue eliminada</span><br/>
	<br/>
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>		

	
<%
	if doWhat="4"then 
	'WhastAPP_Modificar
	%>
 
	<h1>Hoja De Ruta - Modificar Datos</h1>

	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Mostrar_Faltantes.asp" --> 
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Mostrar_Faltantes_Excluidos.asp" --> 
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 

	 
 <%
	End If
%>		

<%
	if doWhat="5"then 
	'WhatsAPP_View_Missing
	%>
 
	<h1>Hoja de Ruta - Agregar Datos</h1>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_GetRowMissing.asp" -->
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Edit_Missing.asp" --> 
	<input type="button" class="btn1" value="Guardar" onClick="chkForm4(this.form,6)" style="width:80px;"> 
	 
 <%
	End If
%>		

<%
	if doWhat="6" then 
		'WhastAPP_Guardar
		hasError=0
%>
		<h1>Hoja de Ruta - Guardar Telefono</h1>
	
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
		<br/><br/>
		<!--#include file="./includes/WhatsAPP_HojaDeRuta_Save_Missing.asp" --> 
<%
		'response.write sSQL
		if hasError=0 then
%>
		<span class="textSuccess">Los cambios fueron guardados</span><br/> 	 
<%
		Else
%>
		<span class="textFail">Error: <%=ErrorMessage%> </span><br/>
		
<%		
		End If
%>
		<br/>
		<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 		
	 
<%
	End If
%>		

<%
	if doWhat="8"then 
	'Whatsapp enviar
	%>
		<h1>Hoja de Ruta - Preparar para envío</h1>
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
		<br/><br/>
		
	<!--#include file="./includes/WhatsAPP_HojaDeRuta_Preparar.asp" --> 
<%
		if hasError=0 then
%>
			<span class="textSuccess"><%=Mensaje%></span><br/> 	 
<%
		Else
%>
			<span class="textFail"><%=Mensaje%> </span><br/>
		
<%		
		End If
%> 
		<br/>
		<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 		
	 
<%
	End If
%>		
 

<%
	if doWhat="9"then 
	'Eliminar item de hoja de ruta  
	%>
		<h1>Hoja de Ruta - Exluir Item</h1>
			
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
		<br/><br/>
			
	<!--#include file="./includes/WhatsAPP_HojaDeRutaDatos_Excluir.asp" --> 
<%
		if hasError=0 then
%>
			<span class="textSuccess"><%=Mensaje%></span><br/> 	 
<%
		Else
%>
			<span class="textFail">Error: <%=Mensaje%> </span><br/>
		
<%		
		End If
%>
		<br/>
		<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 		
	 
	 
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
 
function chkForm2(Fm,prm,param1){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.submit(); 
	}
}
function chkFormExcluir(Fm,prm,param1,param2){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.hojaderutarow.value=param2;
		Fm.excluir.value=1;
		Fm.submit(); 
	}
}
function chkFormIncluir(Fm,prm,param1,param2){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.hojaderutarow.value=param2;
		Fm.excluir.value=0;
		Fm.submit(); 
	}
}

function chkForm3(Fm,prm,param1,param2){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.hojaderutarow.value=param2;
		Fm.submit(); 
	}
}

function chkFormXls(Fm,prm,param1,param2,param3){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.sSQL.value=param2;
		Fm.ReportName.value=param3;
		Fm.submit(); 
	}
}

 
function chkForm4(Fm,prm){
	var msg  = "";
	var question  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;


	if(prm == 6){  
		var input = document.getElementsByName("telefonoid")[0].value; 
	 
		if ((input) == ""){
			msg = "Debe ingresar telefono"; 
			document.getElementById("telefonoiderrormsg").innerHTML = msg;
		}
		else {
			
			var pattern = /^[0-9]+$/;  
			if ((input=="" || (input.length<10))) {     
			  msg = "Debe ingresar un telefono válido";
			  document.getElementById("telefonoiderrormsg").innerHTML = msg;
			}
		}
	} 
	  
	if(msg != "") alert(msg);
	else{ 
	 
		Fm.doWhat.value = prm;
		Fm.submit(); 
	}
}
</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


