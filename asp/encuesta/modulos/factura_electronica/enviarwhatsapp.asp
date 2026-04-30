<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="../../includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="../../includes/db_con_open_WhatsAppapi.asp" --><%
dbCon.CommandTimeout = 0 
 
Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor","exec","execute",_
				  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")

For Each s in Request.Form 
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		Response.Redirect(ErrorPage)
	End If
Next
%><!--#include virtual="../../includes/sql-check.asp"--><%

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

if doWhat="" or doWhat<="3" Then
	Session("FileUploaded") = ""
End If

If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../../login.asp"
End If
 
If doWhat = "-1" Then 
	response.redirect "../../menu.asp"
End If 

Set objRequest = Nothing
%>
<!--#include file="header.asp" -->

	<div class="modulo-content">
		<%--#include file="sidebar.asp" --%>
		
		<div style="flex: 1; padding: 20px;">

			<form  action="../../includes/app_excel_whatsapp.asp" method="post" name="FEXCEL"> 
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
			 
			<%
				if doWhat="" or doWhat="0" then 
			%>
				<!--#include file="../../includes/WhatsAPP_Estado_ServicioFTP.asp" --> 
				<div class="content-card">
					<h2>Estado del Servicio FTP</h2>
					<%if serviceshttp=1 THEN %>
						<div style="color: green; font-weight: bold;">✓ Servicio FTP Activo</div>
					<%else %>
						<div style="color: red; font-weight: bold;">✗ Servicio FTP Inactivo</div>
					<%End If %>
				</div>

				<!--#include file="../../includes/WhatsAPP_HojaDeRuta_EnProceso.asp" --> 
				<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Procesados.asp" --> 

			<%
				End If
			%>	
			
			<%
				if doWhat="1" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,0)">
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Mostrar.asp" --> 
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Mostrar_Excluidos.asp" --> 
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>	

			<%
				if doWhat="2" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Enviar</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,0)">
					<br/><br/>
					<!--#include file="../../includes/WhatsApp_HojaDeRuta_Enviar.asp" --> 
					<div style="color: green; margin-top: 10px;">✓ Hoja de ruta <%=HojaDeRutaID%> ingresada para procesar</div>
					<br />
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>	
			
			<%
				if doWhat="3" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Eliminar</h2>
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Eliminar.asp" --> 
					<input type="button" class="btn btn-danger" value="Volver" onClick="chkForm(this.form,0)">
					<br /><br />
					<div style="color: green;">✓ Hoja de ruta <%=HojaDeRutaID%> eliminada</div>
					<br/>
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>		

			<%
				if doWhat="4" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Modificar Datos</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,0)">
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Mostrar_Faltantes.asp" --> 
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Mostrar_Faltantes_Excluidos.asp" --> 
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>		

			<%
				if doWhat="5" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Agregar Datos</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,4)">
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_GetRowMissing.asp" -->
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Edit_Missing.asp" --> 
					<input type="button" class="btn btn-success" value="Guardar" onClick="chkForm4(this.form,6)">
				</div>
			<%
				End If
			%>		

			<%
				if doWhat="6" then 
					hasError=0
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Guardar Teléfono</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,4)">
					<br/><br/>
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Save_Missing.asp" --> 
					<%
						if hasError=0 then
					%>
						<div style="color: green; margin-top: 10px;">✓ Los cambios fueron guardados</div>
					<%
						Else
					%>
						<div style="color: red; margin-top: 10px;">✗ Error: <%=ErrorMessage%></div>
					<%		
						End If
					%>
					<br/>
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>		

			<%
				if doWhat="8" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Preparar para Envío</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,0)">
					<br/><br/>
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Preparar.asp" --> 
					<%
						if hasError=0 then
					%>
						<div style="color: green; margin-top: 10px;">✓ <%=Mensaje%></div>
					<%
						Else
					%>
						<div style="color: red; margin-top: 10px;">✗ <%=Mensaje%></div>
					<%		
						End If
					%> 
					<br/>
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>		
			 

			<%
				if doWhat="9" then 
			%>
				<div class="content-card">
					<h2>Hoja de Ruta - Eliminar Item</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,4)">
					<br/><br/>
					<!--#include file="../../includes/WhatsAPP_HojaDeRuta_Mostrar_Faltantes_ConChecks.asp" --> 
					<input type="button" class="btn btn-danger" value="Eliminar" onClick="chkForm3(this.form,10,document.FF.HojaDeRutaID.value, document.FF.hojaderutarow.value)">
					<br/>
					<input type="button" class="btn btn-primary" value="Volver al Menú" onClick="volvermenu(this.form)">
				</div>
			<%
				End If
			%>

			</form>

		</div>
	</div>

	</div> <!-- Cierre de modulo-container -->

<script language="javascript">
function volvermenu(Fm){
	Fm.doWhat.value = -1;
	Fm.submit();
}

function chkForm(Fm,prm){
	var msg = "";	
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
			msg = "Debe ingresar teléfono"; 
			document.getElementById("telefonoiderrormsg").innerHTML = msg;
		}
		else {
			var pattern = /^[0-9]+$/;  
			if ((input=="" || (input.length<10))) {     
			  msg = "Debe ingresar un teléfono válido";
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
</html>
<%dbCon.Close
Set dbCon = Nothing%> 
