<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
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

' OJO REQUEST.FORM se puede usar solo una vez por el tema de la subida de files
' NO volver a usar Request.Form despues de estas lineas de abajo
If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If

' Set objRequest = Request.Form
doWhat = objRequest("doWhat") 
apellido = replace(objRequest("apellido"),"'","")
nombre = replace(objRequest("nombre"),"'","")
esLogistica  = objRequest("esLogistica")& ""
esTransportista  = objRequest("esTransportista")& ""
If esLogistica = "" Then
    esLogistica = "0"
End If
If esTransportista = "" Then
    esTransportista = "0"
End If

If doWhat = "-1"  Then 
	response.redirect "../menu.asp"
End If

submit_logout = objRequest("submit_logout")
If doWhat = "-2"  Then 
	submit_logout = "Salir"
	Session("currentUser") = "" 
End If


If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If
volver = objRequest("volver")
If volver = "Menu"  Then 
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
	<form name="FF" method="post" action="altanomina.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
 
  
    <div class="wrap" style="width: 70%" align="center">

	<div style="overflow-x: auto;">
		<h1>ALTA - USUARIO NOMINA</h1>
		
<%
	if  doWhat="" or doWhat="0"  then 
%>	 	
		
			<div>
			<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
			<tr>
				<td>
				<span>Apellido*</span>
				</td>
				
				<td>
					<input type="text" id="apellido" name="apellido" placeholder="apellido" minlength="50" required value="<%=apellido%>"><span class="labelError" id="apellidoerrormsg"></span>
				</td>
			
			</tr>
			<tr>
				<td>
				<span>Nombre*</span>
				</td>
				
				<td>
					<input type="text" id="nombre" name="nombre" placeholder="nombre" minlength="50" required value="<%=nombre%>"><span class="labelError" id="nombreerrormsg"></span>
				</td>
				<td>
				</td>
			</tr>
			<tr>
				<td colspan="2"> <input type="checkbox" id="esLogistica" name="esLogistica" value="1" checked>Logisica</td>
			</tr>			
			<tr>
				<td colspan="2"> <input type="checkbox" id="esTransportista" name="esTransportista" value="1" checked>Transportista</td> 
			</tr>
			<tr>
			<td>
				<input type="button" class="btn1" value="Guardar" onClick="chkForm(this.form, 1)" style="width:80px;">
				</td>
				<td><input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;"></td>	
			</tr>
			</table>
		</div>
		 

<%
	End If 
	 
	 

	if doWhat="1" Then
	' response.write "texto" & txtresol
%>	 

	<!--#include file="./includes/Nomina_Alta_Save.asp" --> 
		 
	  
	
	<div align="center" class="textSuccess">Alta de usuario correcta <br/><input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,'')"></div>
	
	
<%
 	End If				
%>				
		</div>
    </div>
	
	</Form>

	
<script language="javascript">
function fnsalir(Fm){
	Fm.doWhat.value = -2; 
	Fm.submit(); 
} 

function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
} 
function getParameterByName(name, url = window.location.href) {
	name = name.replace(/[\[\]]/g, '\\$&');
    var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, ' '));
}
  

function chkForm(Fm,prm){
	var msg  = "";
	var txt = ""; 
 	if(prm == 1){  
		var input = document.getElementsByName("apellido")[0].value;
		var cliente = document.getElementsByName("nombre")[0].value;
	 
		//var pattern = /^[a-z]+$/;  
		var pattern = /^[^0-9]+$/;
		if ((input) == "" && cliente== ""){
			msg = "Debe ingresar Apellido y Nombre";
			document.getElementById("apellidoerrormsg").innerHTML = msg;
			document.getElementById("nombreerrormsg").innerHTML = msg;
		}
		else {
			
			if(!input.match(pattern) && input!="") {     
			  msg = "Debe ingresar un apellido válido";
			  document.getElementById("apellidoerrormsg").innerHTML = msg;
			}
			if ((msg) == ""){ 
				var str = document.getElementById("nombre").value;
				if(  (!str.match(pattern) && str!="") || (str=="") ){
					msg = "Debe ingresar un nombre válido";
					document.getElementById("nombreerrormsg").innerHTML = msg;
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
function resetForm(Fm) {
	Fm.filesession = "";
	Fm.idReclamo.value = "";
	Fm.doWhat.value = "";
	Fm.submit(); 
}   
	 
 

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 
