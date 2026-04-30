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
 
if doWhat>="0" Then
	hdrid = clng("0"&objRequest("hdrid"))
	nrocliente = objRequest("nrocliente") 
	nrofactura = objRequest("nrofactura") 
	hdridfirst =  clng("0"&objRequest("hdridfirst") )
	nroclientefirst =  clng("0"&objRequest("nroclientefirst") ) 
	hojaderutarow =  clng("0"&objRequest("hojaderutarow") )
	agrupacion =  clng("0"&objRequest("agrupacion") )
	telefonoid =  objRequest("telefonoid") 
End If

If  Session("currentUser") = "" Then
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

			<form name="FF" method="post" action="reenviowhatsapp.asp">
			
			<input type="hidden" name="doWhat" value="<%=doWhat%>">
			<input type="hidden" name="hojaderutarow" value="<%=hojaderutarow%>">
			<input type="hidden" name="agrupacion" value="<%=agrupacion%>">
			
			<%
				if doWhat="" then 
			%>	 	 
				<div class="content-card">
					<h2>Reenvío WhatsApp</h2>
					<div class="form-group">
						<label for="hdrid">Hoja de Ruta*</label>
						<input type="text" id="hdrid" name="hdrid" placeholder="60285" minlength="6" required value="<%=hdrid%>"><span class="labelError" id="hdriderrormsg"></span>
					</div>

					<div class="form-group">
						<label for="nrocliente">Nro Cliente*</label>
						<input type="text" id="nrocliente" name="nrocliente" placeholder="1245" minlength="7" required value="<%=nrocliente%>"><span class="labelError" id="nroclienteerrormsg"></span>
					</div>

					<div class="form-group">
						<label for="nrofactura">Nro Factura*</label>
						<input type="text" id="nrofactura" name="nrofactura" placeholder="1245" minlength="7" required value="<%=nrofactura%>"><span class="labelError" id="nrofacturaerrormsg"></span>
					</div>

					<button type="button" class="btn btn-primary" onclick="chkForm(this.form, 1)">Buscar</button>
				</div>
			<%
				End If 
				
				if doWhat="1" then 
			%>
				<div class="content-card">
					<h2>Resultados de Búsqueda</h2>
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,'')">

					<input type="hidden" name="nrofactura" id="nrofactura" value="<%=nrofactura%>">
					<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">
					<input type="hidden" name="hdridfirst" value="<%=hdrid%>">
					<input type="hidden" name="nroclientefirst" value="<%=nrocliente%>">
					<input type="hidden" name="hdrid" value="<%=hdrid%>">
					
					<!--#include file="../../includes/WhatsAPP_Reenvio_Buscar.asp" -->
					<!--#include file="../../includes/WhatsAPP_Reenvio_Procesados.asp" -->

					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,'')">
				</div>
			<%
				end if 
				
				if doWhat="2" then 
			%> 
				<div class="content-card">
					<h2>Editar Datos de Reenvío</h2>
					<input type="hidden" name="nrofactura" id="nrofactura" value="<%=nrofactura%>">
					<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">
					<input type="hidden" name="hdridfirst" value="<%=hdrid%>">
					<input type="hidden" name="nroclientefirst" value="<%=nrocliente%>">
					<input type="hidden" name="hdrid" value="<%=hdrid%>">
					
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,1)" style="width:80px;">
					<!--#include file="../../includes/WhatsAPP_Reenvio_GetRowMissing.asp" -->  
					<!--#include file="../../includes/WhatsAPP_Reenvio_Edit_Missing.asp" -->

					<input type="button" class="btn btn-success" value="Guardar" onClick="chkForm4(this.form,3)" style="width:80px;"> 
				</div>
			<%
				end if 

				if doWhat="3" then 
					hasError=0
			%> 
				<div class="content-card">
					<h2>Guardar Cambios</h2>
					<input type="hidden" name="nrofactura" id="nrofactura" value="<%=nrofactura%>">
					<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">
					<input type="hidden" name="hdridfirst" value="<%=hdrid%>">
					<input type="hidden" name="nroclientefirst" value="<%=nrocliente%>">
					<input type="hidden" name="hdrid" value="<%=hdrid%>">
					
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,2)" style="width:80px;">
					<br/><br/>
					<!--#include file="../../includes/WhatsAPP_Reenvio_Save_Missing.asp" --> 
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
				if doWhat="4" then 
					hasError=0
			%> 
				<div class="content-card">
					<h2>Actualizar Estado</h2>
					<input type="hidden" name="nrofactura" id="nrofactura" value="<%=nrofactura%>">
					<input type="hidden" name="nrocliente" id="nrocliente" value="<%=nrocliente%>">
					<input type="hidden" name="hdridfirst" value="<%=hdrid%>">
					<input type="hidden" name="nroclientefirst" value="<%=nrocliente%>">
					<input type="hidden" name="hdrid" value="<%=hdrid%>"> 
					
					<input type="button" class="btn btn-primary" value="Volver" onClick="chkForm(this.form,1)" style="width:80px;">
					<br/><br/>
					<!--#include file="../../includes/WhatsAPP_Reenvio_Update.asp" --> 
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

			</form>

		</div>
	</div>

	</div> <!-- Cierre de modulo-container -->
	
<script language="javascript">
function chkControlar(Fm,prm,param1,param2){ 
	Fm.doWhat.value = prm;
	Fm.hdrid.value=param1;
	Fm.nrocliente.value=param2;
	Fm.submit(); 
}

function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
} 

function chkForm4(Fm,prm){
	var msg  = "";

	if(prm == 3){  
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

function chkForm3(Fm,prm,param1,param2,param3){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.hdrid.value=param1;
		Fm.hojaderutarow.value=param2;
		Fm.agrupacion.value=param3;
		Fm.submit(); 
	}
}

function chkForm(Fm,prm){
	var msg  = "";
	var pattern = /^[0-9]+$/;  

	if(prm == 1){  
		var input = document.getElementsByName("hdrid")[0].value;
		var cliente = document.getElementsByName("nrocliente")[0].value;
		var factura = document.getElementsByName("nrofactura")[0].value;
	 
		if ((input) == "" && cliente== ""){
			msg = "Debe ingresar o HDR o Cliente";
			document.getElementById("hdriderrormsg").innerHTML = msg;
			document.getElementById("nroclienteerrormsg").innerHTML = msg;
		}
		else {
			if(!input.match(pattern) && input!="") {     
			  msg = "Debe ingresar una hoja de ruta válida";
			  document.getElementById("hdriderrormsg").innerHTML = msg;
			}
			if ((msg) == ""){
				var str = document.getElementById("nrocliente").value;
				if(  !str.match(pattern) && str!=""){
					msg = "Debe ingresar un cliente válido";
					document.getElementById("nroclienteerrormsg").innerHTML = msg;
				}
			}
		}
		
		if ((msg) == ""){
			var str = document.getElementById("nrofactura").value;
			if(  !str.match(pattern) && str!=""){
				msg = "Debe ingresar una factura válida";
				document.getElementById("nrofacturaerrormsg").innerHTML = msg;
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
	Fm.idFactura.value = "0";
	Fm.doWhat.value = "";
	Fm.submit(); 
}  	 
</script>	  

</body>
</html>
<%dbCon.Close
Set dbCon = Nothing%>
