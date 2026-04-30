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
FileID = objRequest("FileID") 
OrdenID = objRequest("OrdenID") 
cuit = objRequest("cuitid") 
clientenro = objRequest("clientenro")
vendedornro = objRequest("vendedornro")
facturanro = objRequest("facturanro")

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

'response.write "FileID>" &  FileID
'response.write "OrdenID>" &  OrdenID


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
 

	<form  action="../includes/app_excel_ordenes.asp" method="post" name="FEXCEL"> 
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="FileID" value="<%=FileID%>">    
 	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
 
	</form>
	

	<form name="FF" method="post" action="Ordenes.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="FileID" value="<%=FileID%>"> 
	<input type="hidden" name="OrdenID" value="<%=OrdenID%>"> 
	<input type="hidden" name="cuit" value="<%=cuitid%>"> 
	<input type="hidden" name="razonsocial" value="<%=razonsocialnombre%>"> 
	
	
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
	
 
	
    <div class="wrap" style="width: 95%" align="center">
 <%
	if doWhat="" or doWhat="0" then 
	%>
 
 	<h1>Ordenes</h1>
  
  
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	
	<!--#include file="./includes/Ordenes_EnProceso.asp" --> 
	<!--#include file="./includes/Ordenes_Procesadas.asp" --> 
 
	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
 <%
	End If
%>	
	
<%
	if doWhat="1"then 
	'Ordenes_Mostrar
	%>
	<h1>Ordenes</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Ordenes_Mostrar.asp" --> 
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	

	
<%
	if doWhat="2"then 
	'Ordenes_Enviar
	%>
 
	<h1>Ordenes</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Ordenes_Enviar.asp" --> 
	<br />
	<span class="textSuccess">El FileID: <%=FileID%> fue ingresado para procesar</span><br/>
	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	
	
	
<%
	if doWhat="3"then 
	'Ordenes_Eliminar
	%>
 
	<h1>Ordenes</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Ordenes_Eliminar.asp" --> 
	<br />
	<span class="textSuccess">El FileID: <%=FileID%> fue eliminado</span><br/>
	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>		


<%
	if doWhat="11"then 
	'Ordenes_Item_Eliminar
	%>
 
	<h1>Ordenes - Eliminar item</h1>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Ordenes_Item_Eliminar.asp" --> 
	<br />
	<span class="textSuccess">El FileID: <%=FileID%> fue eliminado</span><br/>
	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>		

	
<%
	if doWhat="4"then 
	'Ordenes_Modificar
	%>
 
	<h1>Ordenes Datos Incompletos</h1>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<!--#include file="./includes/Ordenes_View_Missing.asp" --> 
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 

	 
 <%
	End If
%>		

<%
	if doWhat="5"then 
	'Ordenes_Edit_Missing
	%>
 
	<h1>Agregar Datos Incompletos</h1>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
	<!--#include file="./includes/Ordenes_GetRowMissing.asp" -->
	<!--#include file="./includes/Ordenes_Edit_Missing.asp" --> 
	<input type="button" class="btn1" value="Guardar" onClick="chkForm4(this.form,6)" style="width:80px;"> 
	 
 <%
	End If
%>		

<%
	if doWhat="6" then 
		'Ordenes_Guardar
		hasError=0
%>
	<h1>Ordenes</h1>
	<!--#include file="./includes/Ordenes_Save_Missing.asp" --> 
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,4)" style="width:80px;">
		<br/><br/>
<%
		'response.write sSQL
		if hasError=0 then
%>
		<span class="textSuccess">Los cambios del FileID: <%=FileID%> fueron guardados</span><br/> 	 
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
	'Aplicar a transferencia final
	%>
	<h1>Ordenes - Cobranza</h1>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<BR/>
	<BR/>

	<!--#include file="./includes/Ordenes_Mover_A_Cobranzas.asp" --> 
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
	<BR/>
		<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">	 		
	 
<%
	End If
%>

<%
	if doWhat="10"then 
	'Marcar para enviar por correo
	%>
	<h1>Ordenes - Enviar Email</h1>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<BR/>
	<BR/>

	<!--#include file="./includes/Ordenes_Marcar_Para_Envio_Email.asp" --> 
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
	<BR/>
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
		Fm.FileID.value=param1;
		Fm.submit(); 
	}
}

function chkForm3(Fm,prm,param1,param2){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.FileID.value=param1;
		Fm.OrdenID.value=param2;
		Fm.submit(); 
	}
}

function chkFormXls(Fm,prm,param1,param2,param3) {
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.FileID.value=param1;
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
		var cliente = document.getElementsByName("clientenro")[0].value;
		var vendedor = document.getElementsByName("vendedornro")[0].value;
		var factura = document.getElementsByName("facturanro")[0].value;

		if (cliente== "" && vendedor=="" && factura == "" ){
			msg = "Debe ingresar o cliente o vendedor";
			document.getElementById("clientenroerrormsg").innerHTML = msg;
			document.getElementById("vendedornroerrormsg").innerHTML = msg;
			document.getElementById("facturanroerrormsg").innerHTML = msg;
		}
		else { 
			var pattern = /^[0-9]+$/;  
			if ((input=="" || (input.length<8))) {     
			  msg = "Debe ingresar un cuit válido";
			  document.getElementById("cuitiderrormsg").innerHTML = msg;
			}
			if ((factura=="" || (factura.length<5))) {     
			  msg = "Debe ingresar una factura valida";
			  document.getElementById("facturanroerrormsg").innerHTML = msg;
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


