<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_ventas.asp" --><%
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

 
If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If
 
doWhat = objRequest("doWhat")  
PeriodoID = CLng("0" & objRequest("PeriodoID") )
ClienteID = CLng("0" & objRequest("ClienteID") )
VendedorID = objRequest("VendedorID")
VendedorNombre = objRequest("VendedorNombre") 
ClienteNombre = objRequest("ClienteNombre") 
ObjetivoID = CLng("0" & objRequest("ObjetivoID") )
ObjetivoNombre = objRequest("ObjetivoNombre")  
Acction =  CLng("0" & objRequest("Acction") )
txtdescarte = objRequest("txtdescarte")

'response.write doWhat

If doWhat = "-1"  Then 
	response.redirect "../menusuper.asp"
End If

submit_logout = objRequest("submit_logout")
If doWhat = "-2"  Then 
	submit_logout = "Salir"
	Session("currentUser") = "" 
	Session("username") = "" 
End If


If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../menusuper.asp"
End If
volver = objRequest("volver")
If volver = "Menu"  Then 
	response.redirect "../menusuper.asp"
End If

if doWhat="" or doWhat<"3" Then
	Session("FileUploaded") = ""
End If
 
 
SupervisorID = Session("currentUser")
SupervisorNombre = Session("username")

  
  
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

'response.write "<br/>" & PeriodoID
'response.write "<br/>" & doWhat

Set objRequest = Nothing
%>
<HTML>
<HEAD>
<TITLE>ILLANES HNOS SRL</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<link rel="stylesheet" type="text/css" href="../includes/css/bars.css">
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
	<div style="overflow-x: auto;" align="center">
	<br/>

	<span class="pagetitle">Bienvenido: <%=Session("currentUser")%> - <%=Session("username")%> </span> 
 <%
	if doWhat="" or doWhat="0"then 
	%>
	
	
		
	<form name="FF" method="post" action="supervisorescontrol.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">   
	
	<!--#include file="./includes/venPeriodos.asp" -->
	
	</form>

	 
 <%
	End If

	if  doWhat="" or doWhat="0"  then 
%>	 	
 	<form name="FF2" method="post" action="supervisorescontrol.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">  
	<input type="hidden" name="SupervisorID" value="<%=SupervisorID%>">
	<input type="hidden" name="VendedorID" value="<%=VendedorID%>">
	<input type="hidden" name="VendedorNombre" value="<%=VendedorNombre%>">

	<!--#include file="./includes/Supervisores_Vendedores_List.asp" --> 
	<div>
		<br /> 
		<input type="button" class="btn1" value="Menu" onClick="volvermenu(this.form)" style="width:80px;">
		<input type="button" class="btn1" value="Salir" onClick="fnsalir(this.form)" style="width:80px;">
	</div>
	
	</Form>
<%
	End If

	if  doWhat="1"  then 
%>	 	 
 	<form name="FF3" method="post" action="supervisorescontrol.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
	<input type="hidden" name="VendedorID" value="<%=VendedorID%>">
	<input type="hidden" name="VendedorNombre" value="<%=VendedorNombre%>">
	<input type="hidden" name="ClienteID" value="<%=ClienteID%>">
	<input type="hidden" name="ClienteNombre" value="<%=ClienteNombre%>">	

	
	<span  class="pagetitle">Vendedor: <%=VendedorID%> - <%=VendedorNombre%></span> 	 

	<!--#include file="./includes/Supervisores_Vendedor_Cliente_List.asp" --> 
	<div>
		<br />
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	</div>
	
	</Form>
<%  
	End If 
	
	if  doWhat="2"  then 
%>	 	 
 	<form name="FF3" method="post" action="supervisorescontrol.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
	<input type="hidden" name="VendedorID" value="<%=VendedorID%>">
	<input type="hidden" name="VendedorNombre" value="<%=VendedorNombre%>">	
	<input type="hidden" name="ClienteID" value="<%=ClienteID%>">
	<input type="hidden" name="ClienteNombre" value="<%=ClienteNombre%>">
	<input type="hidden" name="Acction" value="<%=Acction%>">	 
	<input type="hidden" name="ObjetivoID" value="<%=ObjetivoID%>">
	<input type="hidden" name="ObjetivoNombre" value="<%=ObjetivoNombre%>">
	
	
	<span  class="pagetitle">Vendedor: <%=VendedorID%> - <%=VendedorNombre%></span> 
	<span  class="pagetitle">Cliente: <%=clienteid%> - <%=ClienteNombre%></span> 	 

	<!--#include file="./includes/Supervisor_Vendedor_Cliente_Objetivos.asp" --> 
	<div>
		<br />
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,1)" style="width:80px;">
	</div>
	
	</Form>
<%  
	End If 


	if  doWhat="3"  then 
	
	
%>	 	
	
 	<FORM name="FF3" method="post" action="supervisorescontrol.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
	<input type="hidden" name="ClienteID" value="<%=ClienteID%>">
	<input type="hidden" name="ClienteNombre" value="<%=ClienteNombre%>">
	<input type="hidden" name="ObjetivoID" value="<%=ObjetivoID%>">
	<input type="hidden" name="ObjetivoNombre" value="<%=ObjetivoNombre%>">
	<input type="hidden" name="VendedorID" value="<%=VendedorID%>">
	<input type="hidden" name="VendedorNombre" value="<%=VendedorNombre%>">	
	<input type="hidden" name="Acction" value="<%=Acction%>">	 

	
	<span  class="pagetitle">Vendedor: <%=VendedorID%> - <%=VendedorNombre%></span> 
	<br/>
	<span  class="pagetitle">Cliente: <%=clienteid%> - <%=ClienteNombre%></span>
	<br/>	
	<span  class="pagetitle">Objetivo:  <%=ObjetivoID%> - <%=ObjetivoNombre%> </span> 	 
		 
	<span  class="pagetitle">Comentarios:</span>
	
	<textarea name="txtdescarte" id="txtdescarte" style="overflow:auto;resize:none" rows="8" cols="20" maxlength="500"></textarea>
	 
	<br/> 
	<br/> 
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,2)" style="width:80px;">
	<input type="button" class="btn1" value="Guardar" onClick="chkForm(this.form,4)" style="width:80px;">
	</Form>	
			
	<%
	
	End If  
 
	if  doWhat="4"  then 
	
	
%>	 	 
	
 	<FORM name="FF4" method="post" action="supervisorescontrol.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
	<input type="hidden" name="ClienteID" value="<%=ClienteID%>">
	<input type="hidden" name="ClienteNombre" value="<%=ClienteNombre%>">
	<input type="hidden" name="ObjetivoID" value="<%=ObjetivoID%>">
	<input type="hidden" name="ObjetivoNombre" value="<%=ObjetivoNombre%>">
	<input type="hidden" name="VendedorID" value="<%=VendedorID%>">
	<input type="hidden" name="VendedorNombre" value="<%=VendedorNombre%>">	
	<input type="hidden" name="Acction" value="<%=Acction%>">	
	<input type="hidden" name="txtdescarte" value="<%=txtdescarte%>">	
	<!--#include file="./includes/Supervisor_Vendedor_Cliente_Objetivos_Guardar.asp" --> 

	
	<span  class="pagetitle">Vendedor: <%=VendedorID%> - <%=VendedorNombre%></span> 
	<br/>
	<span  class="pagetitle">Cliente: <%=clienteid%> - <%=ClienteNombre%></span>
	<br/>	
	<span  class="pagetitle">Objetivo:  <%=ObjetivoID%> - <%=ObjetivoNombre%> </span> 	 
		
	<br/> 
	<br/> 
	<%
	if  Acction=1  then 
	%>	
	<span class="textSuccess">Foto Validad Correctamente </span>
	<%
	else 
	%>
	<span class="textFail">Foto Descartada Correctamente </span>	
	<%
	End If 
	%>	
	
	<br/> 
	<br/> 
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,2)" style="width:80px;">
	</Form>	
			
	<%
	
	End If  

	%>
		
	
	</div>
			 	 
	
<script language="javascript">
function ShowComments(x) {
  alert(x);
}

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

function chkForm3(Fm,prm,param1,param2,param3,param4){ 
	var msg  = ""; 
	if(prm == 2){
		Fm.doWhat.value=prm;
		Fm.ClienteID.value=param1;
		Fm.ClienteNombre.value=param2;
		Fm.ObjetivoID.value = param3;
		Fm.ObjetivoNombre.value=param4; 
		Fm.submit();
	}  
}
function chkForm4(Fm,prm,param1,param2,param3,param4,param5){ 
	var msg  = "";  
	if(prm == 3){ 
		Fm.doWhat.value=prm;  
		Fm.Acction.value=param1;
		Fm.ClienteID.value=param2; 
		Fm.ClienteNombre.value=param3; 
		Fm.ObjetivoID.value = param4; 
		Fm.ObjetivoNombre.value=param5; 
		Fm.submit();
	}  
}
 
function chkForm2(Fm,prm,param1,param2){
	var msg  = ""; 
	if(prm == 1){ 
		Fm.doWhat.value=prm;
		Fm.VendedorID.value=param1;
		Fm.VendedorNombre.value=param2;
	}	
	if(prm == 2){ 
		Fm.doWhat.value=prm;
		Fm.ClienteID.value=param1;
		Fm.ClienteNombre.value=param2;
	}		
	if(msg != "") alert(msg);
	else{  
		Fm.submit(); 
	}
}


function chkForm(Fm,prm){
	var msg  = "";
	var txt = ""; 
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
 

function fun(archivo){
	openWindow(archivo);
	return false;
}
function openWindow(archivo){

if (document.getElementById) {
   w = screen.availWidth;
   h = screen.availHeight;
}  
var w = 340, h = 280;
var popW = 700, popH = 700;

var leftPos = (popW)/2;
var topPos = (popH)/2;



msgWindow = window.open('','popup','width=' + popW + ',height=' + popH + 
                         ',top=' + topPos + ',left=' + leftPos + ',       scrollbars=yes');

msgWindow.document.write 
    ('<HTML><HEAD><meta http-equiv="image/jpg" content="text/html"><TITLE>FACTURA</TITLE></HEAD><BODY><FORM NAME="form1">' +
    '<img width="600" height="600" src="'+archivo+'" ><br/>' +
    '<div align="center"><INPUT TYPE="button" VALUE="CERRAR"onClick="window.close();"></div></FORM></BODY>   </HTML>');
}

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 
