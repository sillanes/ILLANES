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
TransportistaID = objRequest("TransportistaID") 
TransportistaID2 = objRequest("TransportistaID2")  
ZonaID = objRequest("ZonaID") 
' KilosID = Clng("0" & replace(objRequest("KilosID"),",",".")) 
'KilosID = CDbl(Replace(objRequest("KilosID"), ",", "."))
VehiculoID = objRequest("VehiculoID") 
ayudante =  objRequest("tieneayudante")  
chkayudante = objRequest("chkayudante")

tmpKilos = Trim(objRequest("KilosID"))
If tmpKilos = "" Then
    KilosID = 0
Else
    ' Reemplazar punto por coma (adaptar a la config regional del servidor)
    tmpKilos = Replace(tmpKilos, ".", ",")
    If IsNumeric(tmpKilos) Then
        KilosID = CDbl(tmpKilos)
    Else
        KilosID = 0
    End If
End If


'response.write "***** " & ayudante & " **********" 
'response.write "<br/>"
'response.write "***** " & objRequest("chkayudante") & " **********"
'response.write "***** KilosID:" & KilosID & " **********"
If objRequest("chkayudante")  = "on" or ayudante = 1 THEN

	chkayudante = 1 
	ayudante = 1
Else 
	ayudante = 0
	chkayudante = 0
End IF
	 
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
   
	<form name="FF" method="post" action="transportistasasociarhdr.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="HojaDeRutaID" value="<%=HojaDeRutaID%>"> 
	<input type="hidden" name="TransportistaID" value="<%=TransportistaID%>">
	<input type="hidden" name="TransportistaID2" value="<%=TransportistaID2%>"> 
	<input type="hidden" name="tieneayudante" value="<%=ayudante%>"> 
	
	
    <div class="wrap" style="width: 95%" align="center">
	
 	<h1>HOJAS DE RUTA SIN ASOCIAR</h1>
	
 <%
	if doWhat="" or doWhat="0"then 
	%>

  
  
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	
	<!--#include file="./includes/Transportistas_HDR_SinAsociar.asp" -->  
	<!--#include file="./includes/Transportistas_HDR_Asociadas.asp" -->  

	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
 <%
	End If
%>	
	
<%
	if doWhat="1"then 
	'HDR_Mostrar
	%>
	
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<br/>
	<br/>
	
	<div  class="pagetitle">HOJA DE RUTA: <%=HojaDeRutaID%> </div>
	 
	
	<!--#include file="./includes/Transportistas_Asociar_HDR_Transportista.asp" --> 
	
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
	 
 <%
	End If
%>	
<%
if doWhat="2" then 
%>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,1)" style="width:80px;">
	<br/><br/>

	<div class="pagetitle">HOJA DE RUTA: <%=HojaDeRutaID%> </div>

	<!--#include file="./includes/Transportistas_GetRow.asp" -->

	<div class="pagetitle">TRANSPORTISTA: <%=RDNombre%> </div>

	<!-- ? Checkbox para mostrar/ocultar grilla -->
	<label class="pagetitle">

		<input name="chkayudante" type="checkbox" id="chkToggleGrilla"  onchange="toggleGrilla()"> Sale con ayudante
		
		<br/><br/>
			 
		<input id="buttonnext" type="button" class="btn1" value="Siguiente" onClick="chkForm2(document.FF,3,<%=HojaDeRutaID%>,<%=TransportistaID %>,0)" style="width:80px;">
	 
	 
	</label>
	<br/><br/>

	<!-- ? Contenedor que se puede ocultar -->
	<div id="grillaTransportista2" style="display:none;">

		<!--#include file="./includes/Transportistas_Asociar_HDR_Transportista2.asp" --> 
	</div>

	<br/><br/>
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
<%
End If
%>


	
<%
	if doWhat="3"then 
	%>

	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,2)" style="width:80px;">
	<br/>
	<br/>
	
	<div  class="pagetitle">HOJA DE RUTA: <%=HojaDeRutaID%> </div>

	<br/><br/>
	<!--#include file="./includes/Transportistas_GetRow.asp" --> 

	<div  class="pagetitle">TRANSPORTISTA: <%=RDNombre%> </div>
	<div  class="pagetitle">TELEFONO: <%=RDTelefono%> </div>
 
	
	<%If chkayudante = 1 Then%>	 
	<!--#include file="./includes/Transportistas_GetRow2.asp" --> 
	<br/><br/>
	<div  class="pagetitle">AYUDANTE: <%=RDNombre%> </div>
	<div  class="pagetitle">TELEFONO: <%=RDTelefono%> </div>
	<%End If%>
	<!--#include file="./includes/Transportistas_Vehiculo_Sel.asp" --> 
	<!--#include file="./includes/Transportistas_Zona_Sel.asp" --> 
	<!--#include file="./includes/Transportistas_Kilos.asp" --> 
	

	<input type="button" class="btn1" value="Confirmar" onClick="chkForm5(document.FF,4,<%=HojaDeRutaID%>,<%=TransportistaID%>,<%=TransportistaID2%>,<%=chkayudante%>)" style="width:80px;">

	 
 <%
	End If
%>	
	
	
<%
	if doWhat="4"then 
	'guardar cambios
	%>
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	<br /><br />

	<div  class="pagetitle">HOJA DE RUTA: <%=HojaDeRutaID%> </div>

	<!--#include file="./includes/Transportistas_GetRow.asp" --> 

	<div  class="pagetitle">TRANSPORTISTA: <%=RDNombre%> </div>
	<div  class="pagetitle">TELEFONO: <%=RDTelefono%> </div>
	 
	<%If chkayudante= 1 Then%>
	<!--#include file="./includes/Transportistas_GetRow2.asp" --> 
	<br/><br/>
	<div  class="pagetitle">AYUDANTE: <%=RDNombre%> </div>
	<div  class="pagetitle">TELEFONO: <%=RDTelefono%> </div>
	
	
	<%End If%>
	
	<!--#include file="./includes/Transportistas_Vehiculo_GetRow.asp" --> 
	<div  class="pagetitle">VEHICULO: <%=RDPatente%> </div>
	
	<!--#include file="./includes/Transportistas_Zona_GetRow.asp" --> 
	<div  class="pagetitle">ZONA: <%=RDzona%> </div>
	
	<div  class="pagetitle">KILOS: <%=KilosID%> </div>
	
	<!--#include file="./includes/Transportistas_Vincular_HDR_Guardar.asp" --> 
	 <br/> <br/> <br/>

<%
		if hasError=0 then
%>
			<span class="textSuccess"><%=ErrorMessage%></span><br/><br/>
<%
		Else
%> 			<span class="textFail"><%=ErrorMessage%> </span><br/>
<%
		End If
%>		
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
<%
	End If
%>	
	
<%
	if doWhat="5"then 
	'guardar cambios
%>	

	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
	
	<!--#include file="./includes/Transportistas_DesVincular_HDR_Guardar.asp" --> 
	 <br/> <br/> <br/>

<%
		if hasError=0 then
%>
			<span class="textSuccess"><%=ErrorMessage%></span><br/><br/>
<%
		Else
%> 			<span class="textFail"><%=ErrorMessage%> </span><br/>
<%
		End If
%>		
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	
<%	
	end if

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
 
function chkForm2(Fm,prm,param1,param2,param3){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.TransportistaID.value=param2; 
		Fm.tieneayudante.value=param3;
		Fm.submit(); 
	}
}
function chkForm5(Fm,prm,param1,param2,param3,param4){
	var msg  = "";	 
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.HojaDeRutaID.value=param1;
		Fm.TransportistaID.value=param2;
		Fm.TransportistaID2.value=param3; 
		Fm.tieneayudante.value=param4; 
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

function toggleGrilla() {
	var chk = document.getElementById("chkToggleGrilla");
	var div = document.getElementById("grillaTransportista2");
	var div2 = document.getElementById("buttonnext");
	div.style.display = chk.checked ? "block" : "none";
	div2.style.display = chk.checked ? "none" : "block";
}

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


