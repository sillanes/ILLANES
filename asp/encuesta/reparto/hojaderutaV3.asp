<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor","exec","execute",_
                  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
%><!--#include virtual="./includes/sql-check.asp"--><%

For Each s in Request.Form
    If ( CheckStringForSQL(Request.Form(s),"form") ) Then
        PrepareReport("Post Varibale")
        Response.Redirect(ErrorPage)
    End If
Next

Function ToDblSeguro(v)
    If IsNull(v) Or Trim("" & v) = "" Then
        ToDblSeguro = 0
    Else
        ToDblSeguro = CDbl(v)
    End If
End Function

Function NzStr(v)
    If IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Function JsSafe(v)
    Dim t
    t = NzStr(v)
    t = Replace(t, "\", "\\")
    t = Replace(t, "'", "\'")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    JsSafe = t
End Function

Function NumSafe(v)
    Dim t
    t = NzStr(v)
    t = Replace(t, ",", ".")
    NumSafe = t
End Function

Function IntSafe(v)
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or Trim("" & v) = "" Then
        IntSafe = 0
    Else
        IntSafe = CInt(v)
        If Err.Number <> 0 Then
            Err.Clear
            IntSafe = 0
        End If
    End If
    On Error GoTo 0
End Function
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
        .cards-container{display:flex;flex-direction:column;gap:12px}
        .cliente-card{border:1px solid #ddd;border-radius:6px;padding:12px;display:flex;flex-direction:column;box-shadow:0 2px 3px rgba(0,0,0,.05);transition:background .3s}
        .card-row{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap}
        .card-row span{margin-right:8px;font-size:.9rem}
        .card-actions{margin-top:8px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
        .total-bar{margin-top:16px;background:#f7f7f7;border:1px solid #ddd;border-radius:6px;padding:10px;font-weight:bold;display:flex;justify-content:space-between}
        @media(min-width:550px){.cliente-card{flex-direction:row;align-items:center}.card-row{flex:1;margin:0}}

        .estado-entregado{background:#e8f7e4}
        .estado-pendiente{background:#fff9e6}
        .estado-anulada{background:#fdecec}
        .estado-otro{background:#f8f8f8}

        .tooltip{position:relative;display:inline-block}
        .tooltip .tooltiptext{
            visibility:hidden;
            width:220px;
            background-color:#555;
            color:#fff;
            text-align:center;
            padding:6px 8px;
            border-radius:6px;
            position:absolute;
            z-index:1;
            bottom:125%;
            left:50%;
            margin-left:-110px;
            opacity:0;
            transition:opacity .3s
        }
        .tooltip:hover .tooltiptext{visibility:visible;opacity:1}
        .detalle-toggle{cursor:pointer;color:#007bff;text-decoration:underline}

        .search-bar{
          position: sticky;
          top: 0;
          z-index: 50;
          background: #fff;
          border: 1px solid #ddd;
          border-radius: 8px;
          padding: 10px;
          margin: 0 0 14px 0;
          box-shadow: 0 2px 6px rgba(0,0,0,.06);
        }
        .search-row{
          display:flex;
          gap:10px;
          align-items:center;
          flex-wrap:wrap;
        }
        .search-row input{
          flex: 1;
          min-width: 180px;
          padding: 10px 12px;
          border: 1px solid #ccc;
          border-radius: 6px;
          font-size: 0.95rem;
        }
        .search-row button{
          padding: 10px 12px;
          border: 0;
          border-radius: 6px;
          cursor: pointer;
          background:#2e30c7;
          color:#fff;
          font-weight:600;
        }
        .search-row button.secondary{
          background:#6c757d;
        }
        .search-hint{
          font-size: 0.85rem;
          color:#666;
          margin-top:6px;
        }
        .cliente-highlight{
          outline: 3px solid #2e30c7;
          box-shadow: 0 0 0 6px rgba(46,48,199,.12);
          transition: box-shadow .3s, outline .3s;
        }

        .icon-btn-disabled{
          width:36px;
          height:36px;
          border-radius:6px;
          background:#d9d9d9;
          color:#777;
          border:1px solid #c8c8c8;
          display:inline-flex;
          align-items:center;
          justify-content:center;
          cursor:not-allowed;
          pointer-events:none;
          opacity:.95;
          box-sizing:border-box;
        }
        .icon-btn-disabled-large{
          min-height:36px;
          padding:8px 12px;
          border-radius:4px;
          background:#bfc5cc;
          color:#f4f4f4;
          border:1px solid #aab2bb;
          display:inline-flex;
          align-items:center;
          justify-content:center;
          cursor:not-allowed;
          pointer-events:none;
          text-decoration:none;
          box-sizing:border-box;
        }
        .checklist-alert{
          display:flex;
          align-items:center;
          gap:10px;
          background:#fff3cd;
          color:#856404;
          border:1px solid #ffe69c;
          border-radius:8px;
          padding:12px 14px;
          margin-bottom:14px;
        }
        .checklist-alert i{
          font-size:1.2rem;
        }
        .checklist-warning-link{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          width:34px;
          height:34px;
          border-radius:50%;
          background:#fff3cd;
          border:1px solid #e6c65c;
          color:#b77900;
          text-decoration:none;
        }
        .checklist-warning-link:hover{
          background:#ffe8a1;
        }
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

    function limpiarClienteHighlight(){
      var prev = document.querySelector(".cliente-card.cliente-highlight");
      if(prev) prev.classList.remove("cliente-highlight");
    }

    function buscarClienteScroll(){
      var v = (document.getElementById("buscadorClienteId").value || "").trim();
      if(v === ""){
        alert("Ingresá un ClienteID.");
        return;
      }
      if(!/^\d+$/.test(v)){
        alert("El ClienteID debe ser numérico.");
        return;
      }

      var el = document.getElementById("cliente-" + v);
      if(!el){
        alert("No se encontró el cliente " + v + " en esta hoja de ruta.");
        return;
      }

      limpiarClienteHighlight();

      var y = el.getBoundingClientRect().top + window.pageYOffset - 120;
      window.scrollTo({ top: y, behavior: "smooth" });

      el.classList.add("cliente-highlight");
      el.setAttribute("tabindex","-1");
      el.focus({preventScroll:true});

      setTimeout(function(){ el.classList.remove("cliente-highlight"); }, 2500);
    }

    function limpiarBuscadorCliente(){
      document.getElementById("buscadorClienteId").value = "";
      limpiarClienteHighlight();
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
Dim hojaRutaID
hojaRutaID = Request.QueryString("hdrid")

If hojaRutaID = "" Then
  Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas 0," & Session("TransportistaID"))
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

        Dim checkListOKLista, checklistBloqueadoLista
        checkListOKLista = IntSafe(rs("CheckListOK"))
        checklistBloqueadoLista = (checkListOKLista = 0)
%>
  <div class="cliente-card estado-otro">
    <div class="card-row">
		<span><strong>HDR:</strong> <%=rs("HojaDeRutaID")%></span>
		<span><strong>Fecha:</strong> <%=rs("Fecha")%></span>
		<span><strong>Patente:</strong> <%=rs("Patente")%></span>
		<span><strong>Clientes:</strong> <%=rs("TotalClientes")%></span>
        <span><strong>Clientes sin trabajar:</strong> <%=rs("ClientesPendientes")%></span>

        <% If checklistBloqueadoLista Then %>
            <span class="tooltip">
                <a href="https://illanes-encuesta.ddns.net/reparto/vehiculos_home.asp"
                   class="checklist-warning-link"
                   title="Check list del vehículo incompleto">
                   <i class="fas fa-triangle-exclamation"></i>
                </a>
                <span class="tooltiptext">El check list del vehículo del día corriente no está completo.</span>
            </span>
        <% End If %>
    </div>

    <div class="card-actions">
        <% If checklistBloqueadoLista Then %>
            <span class="icon-btn-disabled" title="Debe completar el check list del vehículo">
                <i class="fa-solid fa-truck-fast"></i>
            </span>
        <% Else %>
            <a href="hojaderutaV3.asp?hdrid=<%=rs("HojaDeRutaID")%>" title="Seleccionar hoja de ruta"><i class="fa-solid fa-truck-fast"></i></a>
        <% End If %>

        <% If checklistBloqueadoLista Then %>
            <span class="icon-btn-disabled" title="No disponible hasta completar el check list">
                <i class="fa-solid fa-lock-open"></i>
            </span>
        <% Else %>
            <a href="hojaderutav3_cerrar.asp?hdrid=<%=rs("HojaDeRutaID")%>" title="Cerrar"><i class="fa-solid fa-lock-open"></i></a>
        <% End If %>
    </div>
  </div>
<%
    rs.MoveNext
    Loop
    rs.Close:Set rs=Nothing
    End If
%>
</div>

<% Else %>

<%
Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Cabecera_SelV3 " & hojaRutaID)

Dim checklistBloqueado, checkListOK
checkListOK = 0
checklistBloqueado = False

If Not rs.EOF Then
    checkListOK = IntSafe(rs("CheckListOK"))
    checklistBloqueado = (checkListOK = 0)
    'checklistBloqueado = 0
End If
%>

<div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;flex-wrap:wrap;">
  <h1 style="margin:0;">Hoja de Ruta: <%=hojaRutaID%></h1>

  <a href="hojaderutaV3.asp" class="icon-btn" title="Volver" style="background:#2e30c7;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;margin-left:10px;"><i class="fas fa-arrow-left"></i> Volver</a>

  <% If checklistBloqueado Then %>
      <span class="icon-btn-disabled-large" title="No disponible hasta completar el check list">
         <i class="fas fa-lock"></i>&nbsp;Cerrar
      </span>

      <span class="tooltip">
          <a href="https://illanes-encuesta.ddns.net/reparto/vehiculos_home.asp"
             class="checklist-warning-link"
             title="Check list del vehículo incompleto">
             <i class="fas fa-triangle-exclamation"></i>
          </a>
          <span class="tooltiptext">El check list del vehículo del día corriente no está completo. Click para completarlo.</span>
      </span>
  <% Else %>
      <a href="hojaderutav3_cerrar.asp?hdrid=<%=hojaRutaID%>" class="icon-btn" title="Cerrar" style="background:#27ae60;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;"><i class="fas fa-lock"></i> Cerrar</a>
  <% End If %>
</div>

<% If checklistBloqueado Then %>
<div class="checklist-alert">
    <i class="fas fa-triangle-exclamation"></i>
    <div>
        <strong>Atención:</strong> no puede ingresar, entregar pedidos ni cerrar la hoja de ruta porque el check list del vehículo del día corriente no está completo.
        &nbsp;
        <a href="https://illanes-encuesta.ddns.net/reparto/vehiculos_home.asp" style="font-weight:bold;">
            Ir a completar check list
        </a>
    </div>
</div>
<% End If %>

<% Dim totalACobrar,totalCobrado: totalACobrar = 0 : totalCobrado = 0 %>

<div class="search-bar">
  <div class="search-row">
    <input type="tel"
           id="buscadorClienteId"
           placeholder="🔎 Buscar ClienteID (ej: 67653)"
           inputmode="numeric"
           onkeydown="if(event.key==='Enter'){ buscarClienteScroll(); return false; }">
    <button type="button" onclick="buscarClienteScroll()">Buscar</button>
    <button type="button" class="secondary" onclick="limpiarBuscadorCliente()">Limpiar</button>
  </div>
  <div class="search-hint">Tip: escribí el ID del cliente y presioná Enter.</div>
</div>

<div class="cards-container">
<%
Do Until rs.EOF
   Dim estado, icono, tooltip, statusClass
   estado = IntSafe(rs("estado"))

   Select Case estado
     Case 1
        icono = "<i class='fas fa-check-circle' style='color:green;'></i>"
        tooltip = NzStr(rs("EstadoEntrega"))
        statusClass = "estado-entregado"
     Case 2
        icono = "<i class='fas fa-exclamation-circle' style='color:#e69500;'></i>"
        tooltip = NzStr(rs("EstadoEntrega"))
        statusClass = "estado-anulada"
     Case 3
        icono = "<i class='fas fa-times-circle' style='color:red;'></i>"
        tooltip = NzStr(rs("EstadoEntrega"))
        statusClass = "estado-pendiente"
     Case Else
        icono = "<i class='fas fa-hourglass-half' style='color:gray;'></i>"
        tooltip = NzStr(rs("EstadoEntrega"))
        statusClass = "estado-otro"
   End Select
   
   totalACobrar = totalACobrar + ToDblSeguro(rs("ImporteACobrar"))
   totalCobrado = totalCobrado + ToDblSeguro(rs("TotalCobrado"))
   
   'totalCobrado = totalCobrado + CDbl(0 & rs("TotalCobrado"))

   Dim lat, lon, clienteNombre, direccion, formaPago
   lat = NzStr(rs("Latitud"))
   lon = NzStr(rs("Longitud"))
   clienteNombre = NzStr(rs("ClienteNombre"))
   direccion = NzStr(rs("Direccion"))
   formaPago = NzStr(rs("FormaPago"))
%>

   <div class="cliente-card <%=statusClass%>" id="cliente-<%=rs("ClienteID")%>">
     <div class="card-row">
        <span class="detalle-toggle" onclick="mostrarClienteNombre('<%=JsSafe(clienteNombre)%>','<%=JsSafe(direccion)%>')">
            <strong>Cliente:</strong> <%=rs("ClienteID")%>
        </span>
        <span><strong>Cond:</strong> <%=formaPago%></span>
        <span><strong>Facturas:</strong> <%=rs("TotalFacturas")%></span>
        <span><strong>Imp:</strong> $<%=FormatNumber(ToDblSeguro(rs("ImporteACobrar")),2)%></span>

        <% If UCase(Trim(formaPago)) = "CONTADO" And estado = 1 Then %>
            <span>
                <i class="fas fa-dollar-sign" title="Efectivo"></i> $<%=FormatNumber(CDbl(0 & rs("efectivo")),2)%> &nbsp;
                <i class="fas fa-exchange-alt" title="Transferencia"></i> $<%=FormatNumber(CDbl(0 & rs("transferencia")),2)%>&nbsp;
                <i class="fas fa-money-check" title="Cheque"></i> $<%=FormatNumber(CDbl(0 & rs("cheque")),2)%>
            </span>
        <% Else %>
            <span><strong>Cobrado:</strong> $<%=FormatNumber(CDbl(0 & rs("TotalCobrado")),2)%></span>
        <% End If %>
     </div>

     <div class="card-actions">
        <a href="#"
           class="icon-btn reenviar"
           title="Reenviar al cliente"
           onclick="abrirModalReenvio('<%=hojaRutaID%>','<%=rs("ClienteID")%>');return false;">
           <i class="fas fa-share"></i>
        </a>

        <% If checklistBloqueado Then %>
            <span class="icon-btn-disabled" title="No disponible hasta completar el check list">
               <i class="fa-solid fa-box-open"></i>
            </span>

            <span class="tooltip">
                <a href="https://illanes-encuesta.ddns.net/reparto/vehiculos_home.asp"
                   class="checklist-warning-link"
                   title="Check list del vehículo incompleto">
                   <i class="fas fa-triangle-exclamation"></i>
                </a>
                <span class="tooltiptext">Debe completar el check list del vehículo del día corriente.</span>
            </span>
        <% Else %>
            <a href="entregarV3.asp?clienteid=<%=rs("ClienteID")%>&hdrid=<%=hojaRutaID%>&cond=<%=Server.URLEncode(formaPago)%>"
               class="icon-btn"
               title="Marcar como entregado">
               <i class="fa-solid fa-box-open"></i>
            </a>
        <% End If %>

        <a href="rechazar.asp?clienteid=<%=rs("ClienteID")%>&hdrid=<%=hojaRutaID%>&cond=<%=Server.URLEncode(formaPago)%>" class="icon-btn" title="Rechazar entrega"><i class="fas fa-ban"></i></a>

        <% If estado <> 0 Then %>
           <span class="tooltip"><%=icono%><span class="tooltiptext"><%=Server.HTMLEncode(tooltip)%></span></span>
        <% End If %>

        <% If Trim(lat) <> "" And Trim(lon) <> "" Then %>
           <span><a href="#" title="Ver en mapa" onclick="openMap('<%=NumSafe(lat)%>','<%=NumSafe(lon)%>');return false;"><i class="fas fa-map-marked-alt"></i></a></span>
        <% End If %>

        <span style="position:relative;">
          <a href="hojaderuta_comprobantes.asp?hdrid=<%=hojaRutaID%>&clienteid=<%=rs("ClienteID")%>"
             title="Subir / Ver comprobantes">
             <i class="fas fa-upload"></i>

             <% If IntSafe(rs("CantidadComprobantes")) > 0 Then %>
               <span style="
                 position:absolute;
                 top:-6px;
                 right:-10px;
                 background:#dc3545;
                 color:white;
                 font-size:11px;
                 padding:2px 6px;
                 border-radius:12px;
                 font-weight:bold;">
                 <%=IntSafe(rs("CantidadComprobantes"))%>
               </span>
             <% End If %>
          </a>
        </span>

     </div>
   </div>

<%
rs.MoveNext
Loop
%>
</div>

<%
Set rs = rs.NextRecordset
%>

<div class="total-bar">
    <strong>Resumen</strong>
</div>

<%
Do Until rs.EOF
%>
     <div class="total-bar"><span><strong>Total Facturas</strong> $<%=FormatNumber(CDbl(0 & rs("ImporteACobrar")),2)%></span></div>
     <div class="total-bar"><span><strong>Total Efectivo:</strong> $<%=FormatNumber(CDbl(0 & rs("efectivo")),2)%></span></div>
     <div class="total-bar"><span><strong>Total Cheque:</strong> $<%=FormatNumber(CDbl(0 & rs("cheque")),2)%></span></div>
     <div class="total-bar"><span><strong>Total Transferencia:</strong> $<%=FormatNumber(CDbl(0 & rs("transferencia")),2)%></span></div>
     <div class="total-bar"><span><strong>Total CC:</strong> $<%=FormatNumber(CDbl(0 & rs("CC")),2)%></span></div>
     <div class="total-bar"><span><strong>Total Cobrado</strong> $<%=FormatNumber(CDbl(0 & rs("totalCobrado")),2)%></span></div>
<%
rs.MoveNext
Loop
rs.Close:Set rs=Nothing
%>

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

</body>
</html>