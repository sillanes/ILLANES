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
extracto = objRequest("extracto") 
extractorow = objRequest("extractorow") 
cuit = objRequest("cuitid") 
razonsocial = objRequest("razonsocialnombre") 
StartDate = Replace(objRequest("StartDate"),".","/")
fileupload = objRequest("cmd")
sSQL = objRequest("sSQL") 
nombrearchivo = objRequest("nombrearchivo") 
chksincuit =  objRequest("chksincuit") & ""
If chksincuit="on" Then
	chksincuit=1
Else
	chksincuit=0
End IF


'Response.Write("chksincuit: "&chksincuit& "<br>"&VbCrLf)

'response.write "Extracto>" &  extracto
'response.write "ExtractoRow>" &  extractorow


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
 

	<form  action="../includes/app_excel_adm.asp" method="post" name="FEXCEL"> 
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="extracto" value="<%=extracto%>"> 
	<input type="hidden" name="extractorow" value="<%=extractorow%>"> 
	<input type="hidden" name="cuit" value="<%=cuitid%>"> 
	<input type="hidden" name="razonsocial" value="<%=razonsocialnombre%>"> 
	
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
 
	</form>
	

	<form name="FF" method="post" action="extractos.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="extracto" value="<%=extracto%>"> 
	<input type="hidden" name="extractorow" value="<%=extractorow%>"> 
	<input type="hidden" name="cuit" value="<%=cuitid%>"> 
	<input type="hidden" name="razonsocial" value="<%=razonsocialnombre%>"> 
	
	
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
	
 
	
    <div class="wrap" style="width: 95%" align="center">
 <%
	if doWhat="" or doWhat="0"then 
	%>
 
 	<h1>Extractos</h1>
  
  
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	
	<!--#include file="./includes/Extractos_EnProceso.asp" --> 
	<!--#include file="./includes/Extractos_Procesados.asp" --> 

	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
 <%
	End If
%>	
	
<%
	if doWhat="1"then 
	'Extractos_Mostrar
	%>
	<h1>Extractos</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Extractos_Mostrar.asp" --> 
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	

	
<%
	if doWhat="2"then 
	'Extractos_Enviar
	%>
 
	<h1>Extractos</h1>
	<!--#include file="./includes/Extractos_Enviar.asp" --> 
	<br />
	<span class="textSuccess">El extracto fue ingresado para procesar</span><br/>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	
	
	
<%
	if doWhat="3"then 
	'Extractos_Eliminar
	%>
 
	<h1>Extractos</h1>
	<!--#include file="./includes/Extractos_Eliminar.asp" --> 
	<br />
	<span class="textSuccess">El extracto: <%=extracto%> fue eliminado</span><br/>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>		

<%
	if doWhat="11"then 
	'Extractos_Item_Eliminar
	%>
 
	<h1>Extracto - Eliminar item</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Extractos_Item_Eliminar.asp" --> 
	<br />
	<span class="textSuccess">El extracto: <%=extractorow%> fue eliminado</span><br/>
	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	
	
<%
	if doWhat="4"then 
	'Extractos_Modificar
	%>
 
	<h1>Extrato Datos Incompletos</h1>
	<!--#include file="./includes/Extractos_View_Missing.asp" --> 
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 

	 
 <%
	End If
%>		

<%
	if doWhat="5"then 
	'Extractos_Edit_Missing
	%>
 
	<h1>Agregar Datos Incompletos</h1>
	<!--#include file="./includes/Extractos_GetRowMissing.asp" -->
	<!--#include file="./includes/Extractos_Edit_Missing.asp" --> 
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
	<input type="button" class="btn1" value="Guardar" onClick="chkForm4(this.form,6)" style="width:80px;"> 
	 
 <%
	End If
%>		

<%
	if doWhat="6" then 
		'Extractos_Guardar
		hasError=0
%>
	<h1>Extractos</h1>
	<!--#include file="./includes/Extractos_Save_Missing.asp" --> 
<%
		'response.write sSQL
		if hasError=0 then
%>
		<span class="textSuccess">Los cambios del extracto: <%=extracto%> fueron guardados</span><br/> 	 
<%
		Else
%>
		<span class="textFail">Error: <%=ErrorMessage%> </span><br/>
		
<%		
		End If
%>
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
		<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 		
	 
<%
	End If
%>		

<%
	if doWhat="8"then 
	'Aplicar a transferencia final
	%>
	<!--#include file="./includes/Extractos_Mover_A_Transferencias.asp" --> 
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
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
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
		Fm.extracto.value=param1;
		Fm.submit(); 
	}
}

function chkForm3(Fm,prm,param1,param2){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.extracto.value=param1;
		Fm.extractorow.value=param2;
		Fm.submit(); 
	}
}

function chkFormXls(Fm,prm,param1,param2,param3){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.extracto.value=param1;
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
		var input = document.getElementsByName("cuitid")[0].value;
		var rsocial = document.getElementsByName("razonsocialnombre")[0].value;
		var sincuit = document.getElementsByName("chksincuit")[0].value;
	 
		if ((input) == "" && rsocial== "" && sincuit==0){
			msg = "Debe ingresar o cuit o rsocial";
			document.getElementById("cuitiderrormsg").innerHTML = msg;
			document.getElementById("razonsocialerrormsg").innerHTML = msg;
		}
		else {
			
			var pattern = /^[0-9]+$/;  
			if ((input=="" || (input.length<13)) && sincuit==1) {     
			  msg = "Debe ingresar un cuit válido";
			  document.getElementById("cuitiderrormsg").innerHTML = msg;
			}
			pattern=reg;
			if ((msg) == ""){
				var str = document.getElementById("razonsocial").value;
				if	( str ==""){
					msg = "Debe ingresar una razon social válida";
					document.getElementById("razonsocialerrormsg").innerHTML = msg;
				}
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


