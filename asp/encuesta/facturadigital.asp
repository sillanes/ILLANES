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

'response.write "telefono:" & objRequest("telefonoid")

idCliente = CLng("0" & objRequest("nro"))
idTelefono = objRequest("telefonoid") 
   
   

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
 

	<form name="FF" method="post" action="facturadigital.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">   
    <div class="wrap">
<%
	if doWhat="" or doWhat="0"then 
%>
	
	<div align="center" width="100%">
	
	<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
		<tr>
		<td><img src="../images/reciclado4.png">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</td>
		<td><b>FACTURA DIGITAL</b></td>
		<td>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</td>
		</tr>
	</table> 
	</div>

	<h2><span class="textSuccess"><div style="text-align:center">¡Sumate a la Revolución Verde!</div></span></h2><br/>
	<span class="textSuccess"><div style="text-align:justify">En nuestro compromiso con el medio ambiente, te invitamos a adoptar la factura electrónica. Este sencillo cambio no solo simplifica tu vida al reducir el uso de papel, sino que también contribuye a la conservación de nuestros bosques y a la reducción de residuos. </div></br>
	</span>
	<div align="center">

		<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm" width="100%">
		<tr>
			<td>
			<span>Nro Cuenta:*</span>
			</td>
			<td>
			<input type="text" id="nro" name="nro" placeholder="999999" minlength="6" required value=""><span class="labelError" id="nroerrormsg"></span>
			</td>
		
		</tr>
		<tr>
			<td>
			<span>Telefono*</span>
			</td> 
			<td>
				<input type="text" id="telefonoid" name="telefonoid" placeholder="(299)999-9999" minlength="7" required value=""><span class="labelError" id="telefonoiderrormsg"></span>
			</td>
			<td>
			</td>
		</tr>

		</table>
	</div>

	<br/>
	<div align="center">
		<input type="button" class="btn1" value="Adherirse" onClick="chkForm(this.form,1)" style="width:80px;"> <br/><br/>
		<div style="text-align:justify">A partir del 01/07/2025, usted comenzará a recibir su factura únicamente por WhatsAPP</div><br/>
	</div>
	<br/>
<!--	
	<div align="justify">
	<span class="textSuccess">Al elegir la factura electrónica, estás ayudando a:<br/>


	<ul>
	  <li>Disminuir la tala de árboles: Menos papel significa más árboles en pie</li>
	  <li>Reducir la huella de carbono: Menos transporte y producción de papel implica una menor emisión de gases contaminantes</li>
	  <li>Fomentar la eficiencia: Accede a tus facturas de manera inmediata y segura, desde cualquier lugar</li>
	</ul>
 
	</span>
	</div>
--> 
<%
	End If
%>		

<%
	if doWhat="1" then 
		'WhastAPP_Guardar
		hasError=0
%>
		<h1>FACTURA DIGITAL - Guardar Telefono</h1>
		
		<div align="center">
		<br/><br/>
		<!--#include file="./includes/WhatsAPP_Telefono_Guardar.asp" --> 
<%
		'response.write sSQL
		if hasError=0 then
%>
		<span class="textSuccess">Estimado cliente, gracias por enviarnos su información. Los cambios fueron guardados</span><br/> 	 
<%
		Else
%>
		<span class="textFail">Estimado cliente, hemos tenido un problema. <%=ErrorMessage%> </span><br/>
		
<%		
		End If
%>
		<br/>
		<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
		</div>
<%
	End If
%>		
  
    </div>
	 
	 
	</Form>
	
<script language="javascript">
 
 
function chkForm(Fm,prm){
	var msg  = "";
	var question  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;

	if(prm == 0){  
		 msg="";
	} 
	if(prm == 1){  
		var inputtelefono = document.getElementsByName("telefonoid")[0].value; 
		var inputnro = document.getElementsByName("nro")[0].value; 

		if ((inputnro) == ""){
			msg = "Debe ingresar un nro de cuenta válido"; 
			document.getElementById("nroerrormsg").innerHTML = msg;
		}
		else {
			var pattern = /^[0-9]+$/;  
			if ((inputnro=="" || (inputnro.length<4) ||  !(inputnro.match(pattern)) )) {     
			  msg = "Debe ingresar un nro de cuenta válido";
			  document.getElementById("nroerrormsg").innerHTML = msg;
			} 
		}
		
		if(msg == ""){
		 
			if ((inputtelefono) == ""){
				msg = "Debe ingresar telefono"; 
				document.getElementById("telefonoiderrormsg").innerHTML = msg;
			}
			else {
				
				//var pattern = /^[0-9]+$/;  
				var pattern = /^\(?(\d{3})\)?[- ]?(\d{3})[- ]?(\d{4})$/
				if ((inputtelefono=="" || (inputtelefono.length<7) || !(inputtelefono.match(pattern)) ))  {     
				  msg = "Debe ingresar un telefono válido";
				  document.getElementById("telefonoiderrormsg").innerHTML = msg;
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


