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

idCliente = CLng("0" & objRequest("nro"))
idTelefono = objRequest("telefonoid") 

Set objRequest = Nothing
%>
<!--#include file="header.asp" -->

	<div class="modulo-content">
		<%--#include file="sidebar.asp" --%>
		
		<div style="flex: 1; padding: 20px;">

			<form name="FF" method="post" action="facturadigital.asp">
			<input type="hidden" name="doWhat" value="<%=doWhat%>">   

			<%
				if doWhat="" or doWhat="0" then 
			%>
				<div class="content-card">
					<h2 style="text-align: center; color: #27ae60;">🌱 ¡Sumate a la Revolución Verde!</h2>
					
					<p style="text-align: justify; font-size: 15px;">
						En nuestro compromiso con el medio ambiente, te invitamos a adoptar la factura electrónica. 
						Este sencillo cambio no solo simplifica tu vida al reducir el uso de papel, sino que también 
						contribuye a la conservación de nuestros bosques y a la reducción de residuos.
					</p>

					<div class="form-group">
						<label for="nro">Nro Cuenta*</label>
						<input type="text" id="nro" name="nro" placeholder="999999" minlength="6" required value=""><span class="labelError" id="nroerrormsg"></span>
					</div>

					<div class="form-group">
						<label for="telefonoid">Teléfono*</label>
						<input type="text" id="telefonoid" name="telefonoid" placeholder="(299)999-9999" minlength="7" required value=""><span class="labelError" id="telefonoiderrormsg"></span>
					</div>

					<button type="button" class="btn btn-success" onclick="chkForm(this.form,1)" style="width: 120px;">Adherirse</button>
					
					<p style="text-align: center; margin-top: 20px; font-size: 14px;">
						<strong>A partir del 01/07/2025, usted comenzará a recibir su factura únicamente por WhatsApp</strong>
					</p>

					<div style="background-color: #ecf0f1; padding: 15px; border-radius: 4px; margin-top: 20px;">
						<h4 style="margin-top: 0;">Al elegir la factura electrónica, estás ayudando a:</h4>
						<ul style="margin-bottom: 0;">
						  <li>Disminuir la tala de árboles: Menos papel significa más árboles en pie</li>
						  <li>Reducir la huella de carbono: Menos transporte y producción de papel</li>
						  <li>Fomentar la eficiencia: Accede a tus facturas de manera inmediata y segura</li>
						</ul>
					</div>
				</div>
			<%
				End If
			%>		

			<%
				if doWhat="1" then 
					hasError=0
			%>
				<div class="content-card">
					<h2>Factura Digital - Guardar Teléfono</h2>
					
					<div style="margin-bottom: 20px;">
						<!--#include file="../../includes/WhatsAPP_Telefono_Guardar.asp" --> 
						<%
							if hasError=0 then
						%>
							<div style="color: green; font-weight: bold;">✓ Su solicitud fue registrada correctamente. Pronto recibirá un SMS de confirmación</div>
						<%
							Else
						%>
							<div style="color: red; font-weight: bold;">✗ Error: <%=ErrorMessage%></div>
						<%		
							End If
						%>
						<br/>
					</div>

					<button type="button" class="btn btn-primary" onclick="volvermenu(this.form)">Volver al Menú</button>
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
	var msg  = "";

	if(prm == 1){  
		var nro = document.getElementById("nro").value;
		var telefono = document.getElementById("telefonoid").value;
	 
		if ((nro) == ""){
			msg = "Debe ingresar número de cuenta"; 
			document.getElementById("nroerrormsg").innerHTML = msg;
		}
		else {
			var pattern = /^[0-9]+$/;  
			if (!nro.match(pattern)) {     
			  msg = "Debe ingresar un número de cuenta válido";
			  document.getElementById("nroerrormsg").innerHTML = msg;
			}
		}

		if ((msg) == ""){
			if ((telefono) == ""){
				msg = "Debe ingresar teléfono"; 
				document.getElementById("telefonoiderrormsg").innerHTML = msg;
			}
			else {
				var phonePattern = /^[\d\s\-\+\(\)]+$/;
				if (!telefono.match(phonePattern) || telefono.length < 7) {     
				  msg = "Debe ingresar un teléfono válido";
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
</html>
<%dbCon.Close
Set dbCon = Nothing%>
