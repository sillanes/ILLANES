<%@ Language=VBScript %>
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reparto.asp" --><%
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
	empleadoid = clng("0"&objRequest("empleadoid"))
	telefonoid =  objRequest("telefonoid") 
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


	<form name="FF" method="post" action="Transportistas_Telefonos_Modificar.asp">
	
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	
	
    <div class="wrap" style="width: 70%" align="center">
	<div style="overflow-x: auto;">
		<h1>TELEFONOS</h1>
<% 
	if doWhat="" then 
	%>
	
	
	<input type="hidden" name="empleadoid" id="empleadoid" value="<%=empleadoid%>">
	<input type="hidden" name="telefonoid" id="telefonoid" value="<%=telefonoid%>">
	  
	<!--#include file="./includes/Transportistas_Telefonos_Actualizar.asp" -->

		
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">

	<%
	end if 
 	

	if doWhat="1" then 
	%> 
	
	
	<input type="hidden" name="empleadoid" id="empleadoid" value="<%=empleadoid%>"> 
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,'')" style="width:80px;"> 

	<!--#include file="./includes/Transportistas_Telefonos_GetRow.asp" -->  
	<!--#include file="./includes/Transportistas_Telefonos_Edit_Row.asp" -->
	
	<input type="button" class="btn1" value="Guardar" onClick="chkForm2(this.form,2,<%=empleadoid%>)" style="width:80px;"> 
	


	<%
	end if 
 	%>
		
<%
	if doWhat="2" then 
		'WhastAPP_Guardar
		hasError=0
%> 

		
		<input type="hidden" name="empleadoid" id="empleadoid" value="<%=empleadoid%>">
		<input type="hidden" name="telefono" id="telefonoid" value="<%=telefonoid%>">
		
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,1)" style="width:80px;">
		<br/><br/>
		<!--#include file="./includes/Transportistas_Telefonos_Save_Row.asp" --> 
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
		
    </div>
			
    </div>
	
	
	
	</Form>
	
<script language="javascript">

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

function chkForm2(Fm,prm,param1){
	var msg  = "";
	var question  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;


	if(prm == 2){  
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
		Fm.empleadoid.value=param1; 
		Fm.submit(); 
	}
}


function chkForm(Fm,prm){
	var msg  = "";
	var question  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;
 

	  
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


