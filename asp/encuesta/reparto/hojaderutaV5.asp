<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If



load(rel_Folder)
  Server.ScriptTimeout = 3600
  Dim Form: Set Form = New ASPForm 

  Form.UploadID = Request.QueryString("UploadID")
  Form.SizeLimit = 10*1024*1024 '10MB

  Dim HTML, hResult
  Const fsCompletted  = 0
  Const fsSizeLimit   = &HD
  Const fsTimeOut     = &HE
  Const fsError       = &HA
  
  If Form.State > fsError Then 
    If Form.State = fsSizeLimit Then 
      hResult = "El archivo supera el límite permitido (" & Form.SizeLimit/1024 & " KB)."
    ElseIf Form.State = fsTimeOut Then
      hResult = "El tiempo de carga excedió el máximo permitido (" & Form.ReadTimeout & " seg)."
    Else
      hResult = "Error de carga (código " & Form.State & ")."
    End If
    hResult = "<div class='textError'>" & hResult & "</div>"
    Response.Status = "400 Bad request"
  ElseIf Form.State = fsCompletted Then
    Form.Files.Save MapFolderToDisk(rel_Folder)
    hResult = "<div class='textSuccess'>Archivo subido correctamente.</div>"
  ElseIf Request.QueryString("Action") = "Cancel" Then   
    hResult = "<div class='textWarning'>La carga fue cancelada.</div>"  
  End If 

  Dim UploadID
  UploadID = Form.NewUploadID

  HTML = ""
  HTML = HTML & hResult
  HTML = HTML & "<form name='FFUP' method='post' enctype='multipart/form-data' " & _
                 "action='" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&Step=2&CampaniaID=" & CampaniaID) & "' " & _
                 "onsubmit='return ProgressBar();'>"
  HTML = HTML & "<div class='form-group'>"
  HTML = HTML & "  <label for='File1'>Seleccione un archivo Excel:</label>"
  HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required>"
  HTML = HTML & "</div>"
  HTML = HTML & "<div class='form-actions'>"
  HTML = HTML & "  <button type='submit' class='btn-volver'>Subir Archivo</button>"
  HTML = HTML & "  <button type='button' onclick='window.location.href=""campania_email_configurar.asp?CampaniaID=" & CampaniaID & """' class='btn-cancelar'>Cancelar</button>"
  HTML = HTML & "</div>"
  HTML = HTML & "</form>"

  HTML = HTML & "<script>" & vbCrLf
  HTML = HTML & "function ProgressBar(){" & vbCrLf
  HTML = HTML & "  var ProgressURL = 'progress.asp?UploadID=" & UploadID & "';" & vbCrLf
  HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
  HTML = HTML & "  return true;" & vbCrLf
  HTML = HTML & "}" & vbCrLf
  HTML = HTML & "</script>"

  do_Upload = HTML
End Function

 Function do_Upload_PDF(rel_Folder)
  Server.ScriptTimeout = 3600
  Dim Form: Set Form = New ASPForm 

  Form.UploadID = Request.QueryString("UploadID")
  Form.SizeLimit = 10*1024*1024 '10MB

  Dim HTML, hResult
  Const fsCompletted  = 0
  Const fsSizeLimit   = &HD
  Const fsTimeOut     = &HE
  Const fsError       = &HA
  
  If Form.State > fsError Then 
    If Form.State = fsSizeLimit Then 
      hResult = "El archivo supera el límite permitido (" & Form.SizeLimit/1024 & " KB)."
    ElseIf Form.State = fsTimeOut Then
      hResult = "El tiempo de carga excedió el máximo permitido (" & Form.ReadTimeout & " seg)."
    Else
      hResult = "Error de carga (código " & Form.State & ")."
    End If
    hResult = "<div class='textError'>" & hResult & "</div>"
    Response.Status = "400 Bad request"
  ElseIf Form.State = fsCompletted Then
    Form.Files.Save MapFolderToDisk(rel_Folder)
    hResult = "<div class='textSuccess'>Archivo subido correctamente.</div>"
  ElseIf Request.QueryString("Action") = "Cancel" Then   
    hResult = "<div class='textWarning'>La carga fue cancelada.</div>"  
  End If 

  Dim UploadID
  UploadID = Form.NewUploadID

  HTML = ""
  HTML = HTML & hResult
  HTML = HTML & "<form name='FFUP' method='post' enctype='multipart/form-data' " & _
                 "action='" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&Step=4&CampaniaID=" & CampaniaID) & "' " & _
                 "onsubmit='return ProgressBar();'>"
  HTML = HTML & "<div class='form-group'>"
  HTML = HTML & "  <label for='File1'>Seleccione un archivo PDF:</label>"
  HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required>"
  HTML = HTML & "</div>"
  HTML = HTML & "<div class='form-actions'>"
  HTML = HTML & "  <button type='submit' class='btn-volver'>Subir Archivo</button>"
  HTML = HTML & "  <button type='button' onclick='window.location.href=""campania_email_configurar.asp?CampaniaID=" & CampaniaID & """' class='btn-cancelar'>Cancelar</button>"
  HTML = HTML & "</div>"
  HTML = HTML & "</form>"

  HTML = HTML & "<script>" & vbCrLf
  HTML = HTML & "function ProgressBar(){" & vbCrLf
  HTML = HTML & "  var ProgressURL = 'progress.asp?UploadID=" & UploadID & "';" & vbCrLf
  HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
  HTML = HTML & "  return true;" & vbCrLf
  HTML = HTML & "}" & vbCrLf
  HTML = HTML & "</script>"

  do_Upload_PDF = HTML
End Function

  
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

Set objRequest = Nothing


if doWhat="" or doWhat<"3"   Then
	 Session("FileUploaded") = ""
End If

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
 

%><!--#include virtual="./includes/sql-check.asp"--><%

For Each s in Request.Form 
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		' Redirect to an error page
		Response.Redirect(ErrorPage)
	End If
Next

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Transportista</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        /* ----  CARD LAYOUT  ----*/
        .cards-container{display:flex;flex-direction:column;gap:12px}
        .cliente-card{border:1px solid #ddd;border-radius:6px;padding:12px;display:flex;flex-direction:column;box-shadow:0 2px 3px rgba(0,0,0,.05);transition:background .3s}
        .card-row{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap}
        .card-row span{margin-right:8px;font-size:.9rem}
        .card-actions{margin-top:8px;display:flex;gap:8px}
        .total-bar{margin-top:16px;background:#f7f7f7;border:1px solid #ddd;border-radius:6px;padding:10px;font-weight:bold;display:flex;justify-content:space-between}
        @media(min-width:550px){.cliente-card{flex-direction:row;align-items:center}.card-row{flex:1;margin:0}}

        /* estados de entrega */
        .estado-entregado{background:#e8f7e4}
        .estado-pendiente{background:#fff9e6}
        .estado-anulada{background:#fdecec}
        .estado-otro{background:#f8f8f8}

        /* tooltip */
        .tooltip{position:relative;display:inline-block}
        .tooltip .tooltiptext{visibility:hidden;width:160px;background-color:#555;color:#fff;text-align:center;padding:5px 0;border-radius:6px;position:absolute;z-index:1;bottom:125%;left:50%;margin-left:-80px;opacity:0;transition:opacity .3s}
        .tooltip:hover .tooltiptext{visibility:visible;opacity:1}
        .detalle-toggle{cursor:pointer;color:#007bff;text-decoration:underline}
    </style>
<script>
    let reenviarHDR = null;
    let reenviarCliente = null;

    function abrirModalReenvio(hdrid, clienteid){
      reenviarHDR = hdrid;
      reenviarCliente = clienteid;

      document.getElementById("telefonoReenvio").value = "";
      document.getElementById("modalReenvio").style.display = "block";
    }

    function cerrarModalReenvio(){
      document.getElementById("modalReenvio").style.display = "none";
    }

    function confirmarReenvio(){
      var tel = document.getElementById("telefonoReenvio").value.trim();
      if(tel === ""){
        alert("Debe ingresar un número de teléfono.");
        return;
      }
      if(!/^\+?\d{7,15}$/.test(tel)){
        alert("Ingrese un teléfono válido.");
        return;
      }
      location.href =
        "reenvio_guardar.asp?HojaDeRutaID=" + reenviarHDR +
        "&ClienteID=" + reenviarCliente +
        "&Telefono=" + encodeURIComponent(tel);
    }

    function mostrarClienteNombre(n,a){alert("Cliente: "+n + "\nDireccion: "+a)}
    function openMap(lat,lon){if(lat&&lon){window.open("https://www.google.com/maps?q="+lat+","+lon,"_blank");}}
    function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}

    /* >>> NUEVO <<< */
	function abrirUpload(hdrid, clienteid){
	  document.getElementById("uploadForm").action =
		"hojaderuta_upload.asp?hdrid=" + hdrid + "&clienteid=" + clienteid;

	  document.getElementById("modalUpload").style.display = "block";
	}



    function cerrarUpload(){
      document.getElementById("modalUpload").style.display = "none";
    }
</script>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex:1;">👤 <%=Server.HTMLEncode(Session("NombreTransportista"))%></strong>
    <form method="post" action="logout.asp" style="margin:0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>
<div class="main-content">
<% If Request("msg") = "ok" Then %>
<div style="background:#d4edda;color:#155724;padding:10px;margin-bottom:10px;border:1px solid #c3e6cb;">✅ Hoja de ruta cerrada correctamente.</div>
<% ElseIf Request("msg") = "reenvio_ok" Then %>
<div style="background:#d1ecf1;color:#0c5460;padding:10px;margin-bottom:10px;border:1px solid #bee5eb;">📩 Reenvío registrado con éxito.</div>
<% End If %>
<%
Dim hojaRutaID:hojaRutaID=Request.QueryString("hdrid")
If hojaRutaID="" Then
	'response.write "EXEC usp_Transportista_HojaDeRuta_Abiertas 0," & Session("TransportistaID")
  Set rs=conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas 0," & Session("TransportistaID"))
%>
<h2>Hojas de Ruta Abiertas</h2>
<div class="cards-container">
<% 
	If rs.EOF Then
%>
	<div class="cliente-card estado-otro">
    <div class="card-row"><span><strong>No tiene hojas de rutas pendientes</strong></span></div>
	</div>
<%
	Else
 
	Do Until rs.EOF 
%>
  <div class="cliente-card estado-otro">
    <div class="card-row"><span><strong>HDR:</strong> <%=rs("HojaDeRutaID")%></span><span><strong>Clientes:</strong> <%=rs("TotalClientes")%></span><span><strong>Clientes sin trabajar:</strong> <%=rs("ClientesPendientes")%></span></div>
    <div class="card-actions">
        <a href="hojaderutaV3.asp?hdrid=<%=rs("HojaDeRutaID")%>" title="Seleccionar hoja de ruta"><i class="fa-solid fa-truck-fast"></i></a>
        <a href="hojaderutav3_cerrar.asp?hdrid=<%=rs("HojaDeRutaID")%>" title="Cerrar"><i class="fa-solid fa-lock-open"></i></a>
		<!-- <a href="hojaderutav3_cerrar.asp?hdrid=<%=HojaDeRutaID%>" class="icon-btn" title="Cerrar" style="background:#27ae60;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;"><i class="fas fa-lock"></i> Cerrar</a> -->

    </div>
  </div>
<% 
	rs.MoveNext:Loop:rs.Close:Set rs=Nothing 
	End If
%>
</div>
<% Else '--- detalle de HDR --- %>
<% Set rs=conn.Execute("EXEC usp_Transportista_HojaDeRuta_Cabecera_SelV3 " & hojaRutaID) %>
<div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
  <h1 style="margin:0;">Hoja de Ruta: <%=hojaRutaID%></h1>
  <a href="hojaderutaV3.asp" class="icon-btn" title="Volver" style="background:#2e30c7;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;margin-left:10px;"><i class="fas fa-arrow-left"></i> Volver</a>
  <!-- <a href="cerrarhdr.asp?hdrid=<%=hojaRutaID%>" class="icon-btn" title="Cerrar" style="background:#27ae60;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;"><i class="fas fa-lock"></i> Cerrar</a>-->
  <a href="hojaderutav3_cerrar.asp?hdrid=<%=hojaRutaID%>" class="icon-btn" title="Cerrar" style="background:#27ae60;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;"><i class="fas fa-lock"></i> Cerrar</a>

</div>
<% Dim totalACobrar,totalCobrado:totalACobrar=0:totalCobrado=0 %>
<div class="cards-container">
<% Do Until rs.EOF
   Dim estado:estado=CInt(rs("estado"))
   Dim icono,tooltip,statusClass
   Select Case estado
     Case 1:icono="<i class='fas fa-check-circle' style='color:green;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-entregado"
     Case 2:icono="<i class='fas fa-exclamation-circle' style='color:#e69500;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-anulada"
     Case 3:icono="<i class='fas fa-times-circle' style='color:red;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-pendiente"
     Case Else:icono="<i class='fas fa-hourglass-half' style='color:gray;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-otro"
   End Select
   totalACobrar=totalACobrar+CDbl(rs("ImporteACobrar"))
   totalCobrado=totalCobrado+CDbl(rs("TotalCobrado"))
   
   Dim lat,lon:lat="" : lon=""
   On Error Resume Next
   lat = rs("Latitud"):lon = rs("Longitud")
   
%>
   <div class="cliente-card <%=statusClass%>">
     <div class="card-row">
			 
        <span class="detalle-toggle" onclick="mostrarClienteNombre('<%=Replace(rs("ClienteNombre"),"'","\'")%>','<%=Replace(rs("Direccion"),"'","\'")%>')"><strong>Cliente:</strong> <%=rs("ClienteID")%></span>
        <span><strong>Cond:</strong> <%=rs("FormaPago")%></span>
        <span><strong>Facturas:</strong> <%=rs("TotalFacturas")%></span>
        <span><strong>Imp:</strong> $<%=FormatNumber(rs("ImporteACobrar"),2)%></span>

		<% If UCase(Trim(rs("FormaPago"))) = "CONTADO" And estado = 1 Then %>
			<span>
				<i class="fas fa-dollar-sign" title="Efectivo"></i> $<%=FormatNumber(rs("efectivo"),2)%> &nbsp;
				<i class="fas fa-exchange-alt" title="Transferencia"></i> $<%=FormatNumber(rs("transferencia"),2)%>&nbsp; 		
				<i class="fas fa-money-check" title="Cheque"></i> $<%=FormatNumber(rs("cheque"),2)%> 
			</span>
		<% Else %>
				<span><strong>Cobrado:</strong> $<%=FormatNumber(rs("TotalCobrado"),2)%></span>
		<% End If %>
	
     </div>
     <div class="card-actions"> 
		<a href="#"
		   class="icon-btn reenviar"
		   title="Reenviar al cliente"
		   onclick="abrirModalReenvio('<%=hojaRutaID%>','<%=rs("ClienteID")%>');return false;">
		   <i class="fas fa-share"></i>
		</a>

	    <a href="entregarV3.asp?clienteid=<%=rs("ClienteID")%>&hdrid=<%=hojaRutaID%>&cond=<%= Server.URLEncode(rs("FormaPago"))%>" class="icon-btn" title="Marcar como entregado"><i class="fa-solid fa-box-open"></i></a>
	    <a href="rechazar.asp?clienteid=<%=rs("ClienteID")%>&hdrid=<%=hojaRutaID%>&cond=<%= Server.URLEncode(rs("FormaPago"))%>" class="icon-btn" title="Rechazar entrega"><i class="fas fa-ban"></i></a>

        <% If estado<>0 Then %>
           <span class="tooltip"><%=icono%><span class="tooltiptext"><%=tooltip%></span></span>
        <% End If %>

        <% If lat<>"" And lon<>"" Then %>
           <span><a href="#" title="Ver en mapa" onclick="openMap('<%=replace(lat,",",".")%>','<%=replace(lon,",",".")%>');return false;"><i class="fas fa-map-marked-alt"></i></a></span>
        <% End If %>		
		
		<span>
		  <a href="#"
			 title="Subir comprobante"
			 onclick="abrirUpload('<%=hojaRutaID%>','<%=rs("ClienteID")%>');return false;">
			 <i class="fas fa-upload"></i>
		  </a>
		</span>
		

     </div>

   </div>
<% rs.MoveNext:Loop %>
</div>
<%
set rs = rs.NextRecordset
%>
<div class="total-bar"> 
	<strong>Resumen</strong>
</div>
<%
Do Until rs.EOF
%>
     <div class="total-bar"><span><strong>Total Facturas</strong> $<%=FormatNumber(rs("ImporteACobrar"),2)%> </span></div>
     <div class="total-bar"><span><strong>Total Efectivo:</strong> $<%=FormatNumber(rs("efectivo"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total Cheque:</strong> $<%=FormatNumber(rs("cheque"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total Transferencia:</strong> $<%=FormatNumber(rs("transferencia"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total CC:</strong> $<%=FormatNumber(rs("CC"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total Cobrado</strong> $<%=FormatNumber(rs("totalCobrado"),2)%> </span>  </div>

	
<% rs.MoveNext:Loop:rs.Close:Set rs=Nothing %>
 
<% End If %>
</div>
<%
conn.Close:Set conn=Nothing
%>

<div id="modalReenvio"
     style="display:none; position:fixed; inset:0;
            background:rgba(0,0,0,.55);
            z-index:2147483647;">

  <div style="
      max-width:420px;
      margin:20vh auto;
      background:#fff;
      border-radius:10px;
      overflow:hidden;
      box-shadow:0 10px 30px rgba(0,0,0,.25);">

    <div style="padding:14px 16px; font-weight:700; border-bottom:1px solid #eee;">
      📩 Reenviar al cliente
    </div>

    <div style="padding:14px 16px;">
      <label>Número de teléfono</label>
      <input type="tel"
             id="telefonoReenvio"
             class="form-control"
             placeholder="Ej: 5491123456789">
    </div>

    <div style="display:flex; gap:10px; justify-content:flex-end;
                padding:12px 16px; border-top:1px solid #eee;">
      <button type="button"
              class="btn btn-outline-secondary"
              onclick="cerrarModalReenvio()">
        Cancelar
      </button>

      <button type="button"
              class="btn btn-primary"
              onclick="confirmarReenvio()">
        Reenviar
      </button>
    </div>

  </div>
</div>
<div id="modalUpload"
     style="display:none; position:fixed; inset:0;
            background:rgba(0,0,0,.55);
            z-index:2147483647;">

  <div style="
      max-width:420px;
      margin:20vh auto;
      background:#fff;
      border-radius:10px;
      overflow:hidden;
      box-shadow:0 10px 30px rgba(0,0,0,.25);">

    <form id="uploadForm"
          method="post"
          action="hojaderuta_upload.asp"
          enctype="multipart/form-data">

      <div style="padding:14px 16px; font-weight:700; border-bottom:1px solid #eee;">
        📎 Subir comprobante
      </div>

      <div style="padding:14px 16px;">
        <!-- 🔴 ESTO ES CLAVE -->
        <input type="file"
               name="archivo"
               accept=".jpg,.jpeg,.png,.pdf"
               required>
      </div>

      <div style="display:flex; gap:10px; justify-content:flex-end;
                  padding:12px 16px; border-top:1px solid #eee;">
        <button type="button"
                class="btn btn-outline-secondary"
                onclick="cerrarUpload()">
          Cancelar
        </button>

        <button type="submit"
                class="btn btn-primary">
          Subir
        </button>
      </div>

    </form>
  </div>
</div>


</body>
</html>
