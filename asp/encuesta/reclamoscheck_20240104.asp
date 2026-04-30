<%@ Language=VBScript %> 

<!--#INCLUDE FILE="_upload.asp"-->
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
doWhatPre = objRequest("doWhatPre")
idReclamo = replace(objRequest("idReclamo"),"'","")
txtresol  = objRequest("txtresolucion") 
subirarchivo = objRequest("subirarchivo")
fileupload = objRequest("cmd")


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


if doWhat="" or doWhat<="3" Then
	Session("FileUploaded") = ""
End If

'filesession = Session("FileUploaded")
'response.write "doWhat: " & doWhat
'response.write "<br> idReclamo: " & idReclamo
'response.write "<br> txtresol: " & objRequest("txtresolucion")  
'response.write "<br> file: " & Session("FileUploaded")
'response.write "<br> doWhatPre: " & doWhatPre


 
Function do_Upload(rel_Folder)
  Server.ScriptTimeout = 3600
'Create upload form
'Using Huge-ASP file upload
'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm 
  '{b}Set the upload ID for this form.
  'Progress bar window will receive the same ID.
  Form.UploadID = Request.QueryString("UploadID")'{/b}

  Form.SizeLimit = 10*1024*1024 '10MB

  Dim HTML, hResult
  Const fsCompletted  = 0
  Const fsSizeLimit   = &HD
  Const fsTimeOut     = &HE
  Const fsError       = &HA
  
  If Form.State > fsError Then 'Some error state. 
    If Form.State = fsSizeLimit Then 'Data size exceeds limit. 
      hResult = "Upload size (" & Request.TotalBytes/1024 & "B) exceeds limit (" & Form.SizeLimit/1024 & "kB)."
    ElseIf Form.State = fsTimeOut Then 'Request timeout 
      hResult = "Upload time exceeds limit (" & Form.ReadTimeout & "s)."
    Else
      ' hResult = "Another upload problem (code " & Form.State & ")"
    End If
    hResult = "<Font Color=Red>" & hResult & "</Font><br>"
    Response.Status = "400 Bad request"
  ElseIf Form.State = fsCompletted Then 'Completted
    Form.Files.Save MapFolderToDisk(rel_Folder)
    hResult = "<div><Font Color=Green>" & hResult & "</Font><br></div>"
  ElseIf Request.QueryString("Action") = "Cancel" Then   
    hResult = "Upload was cancelled."  
  End If 

  '{b}get an unique upload ID for this upload script and progress bar.
  Dim UploadID, PostURL, Comment
  UploadID = Form.NewUploadID


  HTML = HTML & hResult
  
  HTML = HTML & "<form name=""FFUP"" method=post ENCTYPE=multipart/form-data Action=" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&idReclamo="&idReclamo) & " OnSubmit=""return ProgressBar();"">"
  HTML = HTML & "</br>"
  HTML = HTML & "<Div ID=files>Adjuntar Archivo: <input type=file name=File1></Div>"
  HTML = HTML & "<input type=submit value=""Subir Archivo""><br>"
  
  HTML = HTML & "</Form>"
  HTML = HTML & "<"+"Script>var nfiles = 1;"
  HTML = HTML & "function ProgressBar(){" & vbCrLf
  HTML = HTML & "  var ProgressURL;" & vbCrLf
  HTML = HTML & "  ProgressURL = 'progress.asp?UploadID=" & UploadID & "'" & vbCrLf
  HTML = HTML & "  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200')" & vbCrLf
  HTML = HTML & "  return true;" & vbCrLf
  HTML = HTML & "};" & vbCrLf
  HTML = HTML & "</"+"Script>"
  HTML = HTML & ""
  do_Upload = HTML
End Function
  



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
	<form name="FF" method="post" action="reclamoscheck.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="doWhatPre" value="<%=doWhatPre%>">
	<input type="hidden" name="idReclamo" value="<%=idReclamo%>"> 
	<input type="hidden" name="txtresol" value="<%=txtresol%>"> 
 	<input type="hidden" name="filesession" value="<%=filesession%>"> 
	<input type="hidden" name="vuelvede" value=""> 
 
	
    <div class="wrap" style="width: 70%" align="center">
		 
<%
	if  doWhat="" then 
%>	 	
 
		<h1>RECLAMOS PENDIENTES</h1> 
		<!--#include file="./includes/TMP_Reclamos_Pendientes_Get.asp" -->
		<span class="textInfoBig">Total Pendientes: <%=totalPendinetes%> </span>
		<br/><br/><br/>
		<h1>RECLAMOS EN PROCESO</h1> 
		<!--#include file="./includes/TMP_Reclamos_EnProceso_Get.asp" -->
		<span class="textInfoBig">Total en Proceso: <%=totalenproceso%> </span>		 
	
		<div>
			<br />
			<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
			<input type="button" class="btn1" value="Salir" onClick="fnsalir(this.form)" style="width:80px;">
		</div>
<%
	End If 
	
	if doWhat>="1" and  doWhat<"4"then 
	%>
	<h1>Reclamo: <%=idReclamo%></h1>
	<!--#include file="./includes/TMP_getReclamoDatos.asp" -->
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
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
			<td align="left" width="60%"><%=dbRS("Cliente")%></td>
			<td align="left" ><%=dbRS("FechaFactura")%></td>
			<td><%=dbRS("FacturaNumero")%></td>
		</tr>
				 
					   
	</table>
	
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
		<tr>
			<td colspan="1" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="1" width="50%"><span class="formHeader">Reclamo</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
	</tbody>
	<tr>
	<td width="50%">
	<span  class="pagetitle" >
		<span align="left"> ¿Su pedido llegó cerrado?</span>
		<span align="left" class="textFail"><%=dbRS("PedidoCerrado")%> </span>
 
	</span>
		  
	<span  class="pagetitle">
		<span>¿Faltó algun producto?</span>
		<span align="left" class="textFail"><%=dbRS("FaltoProducto")%> </span>
		<%if dbRS("FaltoProducto") = "Si" Then %>
		<div align="left">
			<span class="textFail"><%=trim(dbRS("FaltoProductoTexto"))%> </span>
		</div> 
		<%End If%> 
	</span> 
	
	<span  class="pagetitle">
		<span>¿Llegó algo en mal estado?</span>
		<span align="left" class="textFail"><%=dbRS("ProductoMalEstado")%> </span>
		<%if dbRS("ProductoMalEstado") = "Si" Then %>
		<div align="left">
			<span class="textFail"><%=trim(dbRS("ProductoMalEstadoTexto"))%> </span>
		</div> 
		<%End If%>	 
	</span>
 
	<span class="pagetitle">
		<span>Comentario del cliente</span> 
		<div align="left">
			<span class="textFail"><%=trim(dbRS("Comentario"))%> </span>
		</div> 
	</span>
	</td>
	</tr>
	</table>
	<%
	NombreCliente = dbRS("NombreCliente")
	%>
	
	
<%
	End If
	If doWhat = "1" Then
%>	
	<!--#include file="./includes/TMP_Reclamos_Mensajes_History.asp" -->
	<!--#include file="./includes/TMP_Reclamos_Mensajes_Actual.asp" -->

<%
	If  NOT dbRS.EOF  Then
		Mensaje  = dbRS("Mensaje")
	else	
		Mensaje = NombreCliente
	End If
	
%>
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
		<tr>
			<td colspan="1" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="1" width="50%"><span class="formHeader">MENSAJE</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
	</tbody>
	<tr>
		
	<td>
		<span  class="pagetitle">
			<span>Comentario hacia el cliente</span> 
			<textarea id="textarea1" name="txtresolucion" style="overflow:auto;resize:none" wrap="hard" rows="12" cols="20" maxlength="500"><%=Mensaje%></textarea>
		</span>
	</td>
	</tr>
	
	</table>
	<div align="center">
		<input type="button" class="btn1" value="Volver" onClick="resetForm(this.form,1)">
		<input type="button" class="btn1" value="Siguiente" onClick="chkForm3(this.form, 2,1,'<%=idReclamo%>')">
		<input type="button" class="btn1" value="Cerrar Reclamo" onClick="chkForm3(this.form, 4,1,'<%=idReclamo%>')">		
	</div>

<%
	End If
	If doWhat="2"  Then
%>	
	<!--#include file="./includes/TMP_Reclamos_Mensajes_Ins.asp" --> 
	<!--#include file="./includes/TMP_Reclamos_Mensajes_Actual.asp" -->

	
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
		<tr>
			<td colspan="1" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="1" width="50%"><span class="formHeader">Mensaje</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
	</tbody>
	<tr>
		
	<td>
		<span  class="pagetitle">
			<span>Mensaje hacia el cliente</span> 
			<div align="left">
				<span  id="txtre" class="textInfo"><%=trim(dbRS("Mensaje"))%> </span>
			</div> 
	
		</span>
	</td>
	</tr>
	
	</table>
	<div align="center">
		<input type="button" class="btn1" value="Volver" onClick="chkForm2(this.form,1,'<%=idReclamo%>')">
		<input type="button" class="btn1" value="Enviar Mensaje" onClick="chkForm3(this.form, 5, 2,'<%=idReclamo%>')">
		<input type="button" class="btn1" value="Agregar Archivo" onClick="chkForm3(this.form, 3, 2,'<%=idReclamo%>')">
		<input type="button" class="btn1" value="Cerrar Reclamo" onClick="chkForm3(this.form, 4, 2, '<%=idReclamo%>')">		
	</div>
 
 
<%	

	End If	
%>
	</form>
<%		
	if doWhat="3" Then
%> 
	<!--#include file="./includes/TMP_Reclamos_Mensajes_Actual.asp" -->
	
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
	<tbody>
		<tr>
			<td colspan="1" height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td align="left" colspan="1" width="50%"><span class="formHeader">Resolucion</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
	</tbody>
	<tr>
		
	<td>
		<span  class="pagetitle">
			<span>Comentario hacia el cliente</span> 
			<div align="left">
				<span class="textInfo"><%=trim(dbRS("Mensaje"))%> </span>
			</div> 
	
		</span>
	</td>
	</tr>
	
	</table> 

		<%=do_Upload("")%>
		<!--#INCLUDE FILE="_common.asp"--> 
		
	<%
		if Session("FileUploaded") <> "" then
	%>
		<span  class="pagetitle">
			<span>Archivo subido correctamente</span> 
			<div align="left">
				<span class="textSuccess"><%=Session("FileUploaded")%> </span>
			</div> 
	
		</span>		
	<%
		Else
	%>

	<%
		End If
	%>
	  

	 
 
	<div align="center">
		<br/>
		<input type="button" class="btn1" value="Volver" onClick="chkForm3(this.form, 1,3,'<%=idReclamo%>')"> 
		<input type="button" class="btn1" value="Cerrar Reclamo" onClick="chkForm3(this.form, 4,3,'<%=idReclamo%>')">		
	</div>
	

<%	
	End If
	
					
	if doWhat="4" Then
	'response.write "texto" & txtresol
	'<!--#include file="./includes/TMP_Reclamos_Mensajes_Ins.asp" --> 
	' 1 -- Salvar mensaje
	' 2 -- Cerrar Reclamo TMP
	' 2 -- Mover a DBO
	' 3 -- Mostrar Datos
	'<!--#include file="./includes/TMP_Reclamo_Close.asp" -->	
	'<!--#include file="./includes/TMP_Reclamo_Save.asp" -->
	'<!--#include file="./includes/TMP_getReclamoDatos.asp" -->
	
%>	
	<form name="FF2" method="post" action="reclamoscheck.asp">
	<%
	
	if doWhatPre="1" Then 
	%>
	<!--#include file="./includes/TMP_Reclamos_Mensajes_Ins.asp" --> 
	<%
	End If
	%>

	<!--#include file="./includes/TMP_Reclamos_Mensajes_UpdateFile.asp" -->
	<!--#include file="./includes/TMP_Reclamos_Mensajes_Close.asp" -->
	<!--#include file="./includes/Reclamos_Mensajes_Close.asp" -->
	<!--#include file="./includes/TMP_getReclamoDatos.asp" -->
	
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	
	<h1>Reclamo: <%=idReclamo%></h1>
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
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
			<td align="left" width="60%"><%=dbRS("Cliente")%></td>
			<td align="left" ><%=dbRS("FechaFactura")%></td>
			<td><%=dbRS("FacturaNumero")%></td>
		</tr>
				 
					   
	</table>
	
	<div align="center" class="textSuccess">El reclamo <%=idReclamo%> ha sido cerrado <br/><input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,'')"></div>
	
	
	</Form>

<%
 	End If				


	if doWhat="5" Then
	'response.write "texto" & txtresol
%>	
	<form name="FF2" method="post" action="reclamoscheck.asp">

	<!--#include file="./includes/Reclamos_Mensajes_Save.asp" -->
	<!--#include file="./includes/Reclamo_Mensaje_Update_Status.asp" -->
		
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	
	<h1>Reclamo: <%=idReclamo%></h1>
 
	
	<div align="center" class="textSuccess">EL mensaje ah sido enviado al reclamo: <%=idReclamo%> <br/><input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,'')"></div>
	
	
	</Form>

<%
 	End If				
%>				
				
    </div>
	 
	
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

function chkForm3(Fm,prm,prmprv,param1){ 
	var msg  = ""; 
	if(prm == 1){
		Fm.doWhat.value = prm; 
		Fm.doWhatPre.value = prmprv;
		Fm.idReclamo.value=param1; 
		Fm.submit();
	}
	if(prm == 2){
		Fm.doWhat.value = prm; 
		Fm.idReclamo.value=param1; 
		Fm.doWhatPre.value = prmprv;
		//Fm.txtresolucion.value= document.getElementsByName('txtresolucion')[0].value;
		Fm.submit();
	}	
	if(prm == 3){
		Fm.doWhat.value = prm; 
		Fm.idReclamo.value=param1; 
		Fm.doWhatPre.value = prmprv;
		Fm.submit();
	}
	if(prm == 4){
		Fm.doWhat.value = prm; 
		Fm.idReclamo.value=param1; 
		Fm.doWhatPre.value = prmprv;
		Fm.submit();
	}
	if(prm == 5){
		Fm.doWhat.value = prm; 
		Fm.idReclamo.value=param1; 
		Fm.doWhatPre.value = prmprv;
		Fm.submit();
	}	
}
 
function chkForm2(Fm,prm,param1){
	var msg  = "";

	if(prm == 1){ 
		Fm.idReclamo.value=param1;
	}		
	if(msg != "") alert(msg);
	else{ 
		Fm.doWhat.value = prm; 
		Fm.idReclamo.value=param1; 
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
	 
 

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 
