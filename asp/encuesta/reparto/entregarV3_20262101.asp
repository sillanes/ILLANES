<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If
%>

<!--#include file="sidebar.asp" -->
<%
Function FormatearParaJS(ByVal valor)
  Dim s
  s = Trim(CStr(valor))
  s = Replace(s, ".", "")
  s = Replace(s, ",", ".")
  FormatearParaJS = s
End Function

Dim hojaRutaID, clienteID, rs, sql, totalCobrar, clienteNombre, totalFacturas, cuentaCorriente

' ================================================
' CONFIG: HABILITAR RESTA AÚN SI TIENE DESCUENTO
' 0 = DESHABILITADO (NO PERMITE RESTAR)
' 1 = HABILITADO  (PERMITE RESTAR IGUALMENTE)
' ================================================
Dim permitirRestarConDto
permitirRestarConDto = 1   ' <-- CAMBIAR A 0 PARA VOLVER A BLOQUEAR


hojaRutaID = Request.QueryString("hdrid")
clienteID = Request.QueryString("clienteid")
cond = Request.QueryString("cond")

If hojaRutaID = "" Or clienteID = "" or cond= "" Then
  Response.Write "Faltan parámetros."
  Response.End
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ILLANES HNOS SRL</title>

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="estilos.css"> 

  <style>
    .formadepago-efectivo{background:#2c9c55}
    .formadepago-cheque{background:#fff9e6}
    .formadepago-transferencia{background:#6974f9}
    .formadepago-otro{background:#1322d7}

    /* ESTILOS PARA FACTURA ANULADA */
    .bg-danger-subtle { background-color: #f8d7da !important; }
    .factura-card.border-danger { border-width: 2px !important; }
  </style> 
</head>

<body class="bg-light">

<header>
    <button class="menu-toggle" onclick="toggleSidebar()">
        <i class="fas fa-bars"></i>
    </button>

    <strong style="flex: 1;">
        👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %>
    </strong>

    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="container my-4">

<form id="formPagos" method="post" action="registrar_pagoV3.asp" class="mb-4">

<div class="d-flex justify-content-between align-items-center mb-4">
    <h2>Pedidos</h2> 
    <a href="hojaderutav3.asp?hdrid=<%= hojaRutaID %>"  
       class="icon-btn" 
       title="Volver"
       style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none; margin-right: 10px;">
        <i class="fas fa-arrow-left"></i> Volver
    </a> 
</div>

<%
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Sel " & hojaRutaID & ", " & clienteID & ", '" & Replace(cond,"'","''") & "'"
Set rs = conn.Execute(sql)

If rs.EOF Then
  Response.Write "No se encontraron datos para este cliente."
  Response.End
End If

clienteID = rs("clienteID") 
clienteNombre = rs("clienteNombre")
totalCobrar = rs("ImporteTotal")
ImporteTotal = rs("ImporteTotal")
totalFacturas = rs("CantidadFacturas")
cuentaCorriente = rs("CuentaCorriente")
If cuentaCorriente = 1 Then totalCobrar = 0.00
%>

<% If cuentaCorriente = 1 Then %>
    <div class="alert alert-info">
      Este cliente utiliza <strong>Cuenta Corriente</strong>. No es necesario ingresar pagos.
    </div>

    <div class="mb-3 mt-3">
      <label for="dniCC" class="form-label">DNI de la persona que recibe</label>
      <input 
          type="number" 
          class="form-control" 
          id="dniCC" 
          name="dni"
          inputmode="numeric"
          pattern="[0-9]{7,10}"
          placeholder="Ingrese DNI (solo números)" 
          required>
    </div>
<% End If %>

<div class="card mb-3">
    <div class="card-body">
      <p><strong>Hoja de Ruta:</strong> <%= hojaRutaID %></p>
      <p><strong>Cliente:</strong> <%= clienteID %> - <%= clienteNombre %></p>
      <p><strong>Facturas:</strong> <%= totalFacturas %></p>
      <p><strong>Total a Cobrar:</strong> $ <span id="totalGeneral"
											  data-totalorig="<%= FormatearParaJS(ImporteTotal) %>">
											<%= FormatNumber(ImporteTotal, 2) %>
										 </span>
	</p>
    </div>
</div>

<button type="button" class="btn btn-outline-secondary mb-3" onclick="restaurarCantidades()">
    Restablecer Cantidades
</button>

<%
Do Until rs.EOF
%>

<div class="card mb-3 factura-card"
     data-facturaid="<%= rs("FacturaID") %>"
     data-totalorig="<%= FormatearParaJS(rs("TotalFacturas")) %>">


    <div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
        <div>
            <%= rs("Factura") %> - 
            Total: $<span class="factura-total" data-facturaid="<%= rs("FacturaID") %>"><%= FormatNumber(rs("TotalFacturas"), 2) %></span>

            <% If rs("isOR") = 1 Then %>
                <span class="badge bg-warning text-dark ms-2">Tiene una orden de Retiro</span>
            <% End If %>
        </div>

        <!-- BOTÓN ANULAR FACTURA COMPLETA -->
		<button type="button"
				class="btn btn-sm btn-danger"
				onclick="abrirModalAnular('<%= rs("FacturaID") %>')">
			<i class="fas fa-times"></i> Anular
		</button>
    </div>

    <ul class="list-group list-group-flush">

    <%
    facturaID = rs("FacturaID")
    sqlItems = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Detalle_Sel " & facturaID
    Set rsItems = conn.Execute(sqlItems)

    Do Until rsItems.EOF
    %>

        <li class="list-group-item d-flex justify-content-between align-items-start item-factura" 
            data-itemid="<%= rsItems("Id") %>" 
            data-facturaid="<%= facturaID %>" 
            data-precio="<%= FormatearParaJS(rsItems("PrecioUnitario")) %>" 
            data-iva="<%= FormatearParaJS(rsItems("ImporteConIVA")) %>" 
            data-iibb="<%= FormatearParaJS(rsItems("PercepcionIIBB")) %>" 
            data-totalimp="<%= FormatearParaJS(rsItems("TotalConImpuestos")) %>" 
            data-cantorig="<%= rsItems("Cantidad") %>">

            <div class="ms-2 me-auto">
              <div class="fw-bold"><%= rsItems("Descripcion") %></div>
              <div>
                Articulo: <%= rsItems("Articulo") %> 
                &nbsp;&nbsp;Cant: <span class="item-cantidad"><%= rsItems("Cantidad") %></span>

                <input type="hidden"
                       name="item_<%= facturaID %>_<%= rsItems("Id") %>"
                       value="<%= rsItems("Cantidad") %>"
                       class="input-cantidad-item">

                <% If FormatNumber(rsItems("Dto"),2) <> 100 Then %>
					<button type="button" class="btn btn-sm btn-outline-secondary ms-1 btn-restar">-</button>
						
					<% If FormatNumber(rsItems("Dto"),2) > 0 Then %>
						<span class="badge bg-warning text-dark ms-2">Dto</span>
					<% End If %>
                
				<% Else %>
                    <span class="badge bg-warning text-dark ms-2">Dto</span>
                <% End If %>
              </div>

              <small>
                Unitario: $<%= FormatNumber(rsItems("PrecioUnitario"), 2) %><br>
                IVA: $<%= FormatNumber(rsItems("ImporteConIVA"), 2) %><br>
                IIBB: $<%= FormatNumber(rsItems("PercepcionIIBB"), 2) %><br>

                <strong>Total c/ Impuestos:
                    $<span class="item-total-imp"><%= FormatNumber(rsItems("TotalConImpuestos"), 2) %></span>
                </strong>
              </small>

            </div>
        </li>

    <%
        rsItems.MoveNext
    Loop

    rsItems.Close
    %>

    </ul>
</div>

<%
rs.MoveNext
Loop
rs.Close
%>

<input type="hidden" name="hdrid" value="<%= hojaRutaID %>">
<input type="hidden" name="clienteid" value="<%= clienteID %>">
<input type="hidden" name="esCC" value="<%= cuentaCorriente %>">
<input type="hidden" name="cond" value="<%= cond %>">

<!-- ===== GEOLOCALIZACIÓN (AGREGADO) ===== -->
<input type="hidden" name="latitud" id="latitud">
<input type="hidden" name="longitud" id="longitud">
<input type="hidden" id="geo_ok" value="0">


<h4 class="mb-3">Formas de Pago</h4>

<div class="row g-2">
<%
Dim formasPago: formasPago = Array("Efectivo", "Cheque/Echeq", "Transferencia")

For i = 0 To UBound(formasPago)

Select Case i
	Case 0:statusClass="formadepago-efectivo"
	Case 1:statusClass="formadepago-cheque2"
	Case 2:statusClass="formadepago-transferencia" 
	Case Else: statusClass=""
End Select

%>

    <div class="col-6 col-md-4 <%=statusClass%>">
      <label for="pago<%= i+1 %>" class="form-label"><%= formasPago(i) %></label>
      <input 
        id="pago<%= i+1 %>" 
        name="pago<%= i+1 %>" 
        type="number" 
        class="form-control" 
        step="0.01" 
        min="0" 
        value="0">
      <input type="hidden" name="codigo<%= i+1 %>" value="<%= i+1 %>">
    </div>

<%
Next
%>
</div>

<div class="mt-3 p-3 bg-white border rounded">
  <h5>
    Saldo restante:
    $<span id="saldoRestante">
        <% If cuentaCorriente=0 Then %>
            <%= FormatNumber(totalCobrar, 2) %>
        <% Else %>
            <%= FormatNumber(0, 2) %>
        <% End If %>
    </span>
  </h5>
</div>

<div class="mt-3">
  <label for="observaciones" class="form-label">Observaciones</label>
  <textarea 
      id="observaciones" 
      name="observaciones" 
      class="form-control" 
      rows="3" 
      placeholder="Escriba aquí alguna observación..."></textarea>
</div>

<div class="d-grid mt-3">
<button type="button"
        id="btnRegistrar"
        class="btn btn-success btn-lg">
    Registrar Pago
</button>
</div>

<div class="d-grid mt-3">

 

</form>
</div>
<script>
function parseFloatLocal(str) {
  if (!str) return 0;
  return parseFloat(str.replace(/\./g, '').replace(',', '.').trim()) || 0;
}

function formatNumberAr(num) {
  return num.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function actualizarTotales() {

  let totalGeneral = 0;

  document.querySelectorAll('.factura-card').forEach(card => {

    const facturaID = card.dataset.facturaid;
    let totalFactura = 0;

    // ¿Está anulada?
    if (card.classList.contains('border-danger')) {
      totalFactura = 0;
    } else {

      const items = card.querySelectorAll('.item-factura');

      if (items.length > 0) {
        // 👉 Factura normal: sumar ítems
        items.forEach(item => {
          const cantActual   = parseInt(item.querySelector('.item-cantidad').textContent);
          const cantOriginal = parseInt(item.dataset.cantorig);
          const totalImpOrig = parseFloat(item.dataset.totalimp);

          if (cantOriginal > 0 && !isNaN(cantActual)) {
            const totalItem = (totalImpOrig / cantOriginal) * cantActual;

            item.querySelector('.item-total-imp').textContent = formatNumberAr(totalItem);
            item.querySelector('.input-cantidad-item').value = cantActual;

            if (cantActual < cantOriginal) {
              item.classList.add('bg-warning-subtle');
            } else {
              item.classList.remove('bg-warning-subtle');
            }

            totalFactura += totalItem;
          }
        });

      } else {
        // 👉 NC o factura sin ítems
        totalFactura = parseFloat(card.dataset.totalorig || 0);
      }
    }

    // Pintar total por factura
    const totalSpan = document.querySelector(
      '.factura-total[data-facturaid="' + facturaID + '"]'
    );

    if (totalSpan) {
      totalSpan.textContent = formatNumberAr(totalFactura);
    }

    totalGeneral += totalFactura;
  });

  // Total general
  if (<%= cuentaCorriente %> == 0) {
    document.getElementById('totalGeneral').textContent = formatNumberAr(totalGeneral);
    actualizarSaldo();
  }
}


function actualizarSaldo() {
  if (<%= cuentaCorriente %> == 1) return;

  const total = parseFloatLocal(document.getElementById('totalGeneral').textContent);

  let sumaPagos = 0;
  for (let i = 1; i <= 3; i++) {
    const val = parseFloat(document.getElementById('pago' + i).value.replace(',', '.'));
    if (!isNaN(val)) sumaPagos += val;
  }

  const saldo = total - sumaPagos;
  document.getElementById('saldoRestante').textContent = formatNumberAr(saldo);
}

function restaurarCantidades() {

  // 1️⃣ Restaurar cantidades de ítems
  document.querySelectorAll('.item-factura').forEach(item => {
    const cantOrig = item.dataset.cantorig;
    const spanCant = item.querySelector('.item-cantidad');
    const inputCant = item.querySelector('.input-cantidad-item');

    if (spanCant) spanCant.textContent = cantOrig;
    if (inputCant) inputCant.value = cantOrig;

    item.classList.remove('bg-warning-subtle', 'bg-danger-subtle');
  });

  // 2️⃣ Restaurar total de cada factura (incluye NC)
  document.querySelectorAll('.factura-card').forEach(card => {
    const totalOrig = parseFloat(card.dataset.totalorig || 0);
    const facturaID = card.dataset.facturaid;

    const totalSpan = document.querySelector(
      '.factura-total[data-facturaid="' + facturaID + '"]'
    );

    if (totalSpan) {
      totalSpan.textContent = formatNumberAr(totalOrig);
    }

    card.classList.remove('border', 'border-danger');
  });

  // 3️⃣ Restaurar total general original
  const totalGeneralEl = document.getElementById('totalGeneral');
  const totalGeneralOrig = parseFloat(totalGeneralEl.dataset.totalorig || 0);

  totalGeneralEl.textContent = formatNumberAr(totalGeneralOrig);

  // 4️⃣ Recalcular saldo
  actualizarSaldo();
}


document.querySelectorAll('.btn-restar').forEach(btn => {
  btn.addEventListener('click', function () {
    const cantSpan = this.closest('li').querySelector('.item-cantidad');
    let actual = parseInt(cantSpan.textContent);
    if (actual > 0) {
      cantSpan.textContent = actual - 1;
      actualizarTotales();
    }
  });
});

 document.getElementById("btnRegistrar").addEventListener("click", function () {


  // ===============================
  // VALIDAR DNI SI ES CUENTA CORRIENTE
  // ===============================
  if (<%= cuentaCorriente %> === 1) {
    var dniEl = document.getElementById("dniCC");

    if (!dniEl || dniEl.value.trim() === "") {
      alert("Debe ingresar el DNI de la persona que recibe.");
      dniEl?.focus();
      return;
    }

    if (!/^\d{7,10}$/.test(dniEl.value.trim())) {
      alert("El DNI debe tener solo números (entre 7 y 10 dígitos).");
      dniEl.focus();
      return;
    }
  }


  // Si ya tengo geo → envío
  if (document.getElementById("geo_ok").value === "1") {
    document.getElementById("formPagos").submit();
    return;
  }

  if (!navigator.geolocation) {
    alert("Este dispositivo no soporta geolocalización.");
    return;
  }

  navigator.geolocation.getCurrentPosition(
    function (pos) {
      document.getElementById("latitud").value  = pos.coords.latitude.toFixed(6);
      document.getElementById("longitud").value = pos.coords.longitude.toFixed(6);
      document.getElementById("geo_ok").value   = "1";

      // 🔥 recién acá se envía el form
      document.getElementById("formPagos").submit();
    },
    function () {
      alert(
        "⚠️ No se pudo obtener la ubicación.\n\n" +
        "Si estás en WhatsApp: tocá ⋮ → Abrir en Chrome."
      );
    },
    {
      enableHighAccuracy: true,
      timeout: 15000,
      maximumAge: 0
    }
  );
});

document.querySelectorAll('input[type="number"]').forEach(input => {
  input.addEventListener('input', actualizarSaldo);
});

// ===== MODAL ANULAR (WhatsApp-safe) =====
var facturaPendiente = null;

function abrirModalAnular(facturaID){
  facturaPendiente = facturaID;
  document.getElementById("modalAnular").style.display = "block";
}

function cerrarModalAnular(){
  facturaPendiente = null;
  document.getElementById("modalAnular").style.display = "none";
}

function confirmarAnulacion(){
  if (!facturaPendiente) return;

  // input hidden
  if (!document.querySelector('input[name="anulada_' + facturaPendiente + '"]')) {
    var h = document.createElement("input");
    h.type  = "hidden";
    h.name  = "anulada_" + facturaPendiente;
    h.value = "1";
    document.getElementById("formPagos").appendChild(h);
  }

  // poner todo en 0 (con checks para no romper JS)
  document.querySelectorAll('.item-factura[data-facturaid="' + facturaPendiente + '"]')
    .forEach(function(item){
      var c = item.querySelector(".item-cantidad");
      var i = item.querySelector(".input-cantidad-item");
      var t = item.querySelector(".item-total-imp");

      if (c) c.textContent = "0";
      if (i) i.value = "0";
      if (t) t.textContent = "0,00";

      item.classList.add("bg-danger-subtle");
    });

  var card = document.querySelector('.factura-card[data-facturaid="' + facturaPendiente + '"]');
  if (card) card.classList.add("border","border-danger");

  if (typeof actualizarTotales === "function") actualizarTotales();

  cerrarModalAnular();
}

// cerrar tocando fondo
document.getElementById("modalAnular").addEventListener("click", function(e){
  if (e.target === this) cerrarModalAnular();
});

// ===== GEOLOCALIZACIÓN =====  
function solicitarGeolocalizacion() {

  if (!("geolocation" in navigator)) {
    alert("Este dispositivo no soporta geolocalización.");
    return;
  }

  navigator.geolocation.getCurrentPosition(
    function (pos) {
      document.getElementById("latitud").value  = pos.coords.latitude.toFixed(6);
      document.getElementById("longitud").value = pos.coords.longitude.toFixed(6);
      document.getElementById("geo_ok").value   = "1";

      // ✅ habilitar botón
      document.getElementById("btnRegistrar").disabled = false;
    },
    function (err) {
      alert(
        "⚠️ Es obligatorio permitir la ubicación para registrar el pago.\n\n" +
        "Si la bloqueaste, abrí el link en Chrome."
      );
      document.getElementById("geo_ok").value = "0";
    },
    {
      enableHighAccuracy: true,
      timeout: 15000,
      maximumAge: 0
    }
  );
   
} 
 
document.getElementById("btnRegistrar").addEventListener("click", function () {

  // Si ya tengo geolocalización → submit directo
  if (document.getElementById("geo_ok").value === "1") {
    document.getElementById("formPagos").submit();
    return;
  }

  // Pedir geolocalización DESDE el botón
  if (!("geolocation" in navigator)) {
    alert("Este dispositivo no soporta geolocalización.");
    return;
  }

  navigator.geolocation.getCurrentPosition(
    function (pos) {
      document.getElementById("latitud").value  = pos.coords.latitude.toFixed(6);
      document.getElementById("longitud").value = pos.coords.longitude.toFixed(6);
      document.getElementById("geo_ok").value   = "1";

      // 🔥 ahora sí, submit real
      document.getElementById("formPagos").submit();
    },
    function (err) {
      alert(
        "⚠️ No se pudo obtener la ubicación.\n\n" +
        "Si estás en WhatsApp, tocá los 3 puntos → Abrir en Chrome."
      );
    },
    {
      enableHighAccuracy: true,
      timeout: 15000,
      maximumAge: 0
    }
  );
}); 

  
</script>



<div id="modalAnular"
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
      Confirmar anulación
    </div>

    <div style="padding:14px 16px;">
      ¿Seguro que querés <b>anular esta factura completa</b>?
    </div>

    <div style="display:flex; gap:10px; justify-content:flex-end;
                padding:12px 16px; border-top:1px solid #eee;">
      <button type="button"
              class="btn btn-outline-secondary"
              onclick="cerrarModalAnular()">
        Cancelar
      </button>

      <button type="button"
              class="btn btn-danger"
              onclick="confirmarAnulacion()">
        Anular
      </button>
    </div>

  </div>
</div>


</body>
</html>

<%
If Not rsItems Is Nothing Then
  If Not rsItems.State = 0 Then rsItems.Close
  Set rsItems = Nothing
End If

If Not rs Is Nothing Then
  If Not rs.State = 0 Then rs.Close
  Set rs = Nothing
End If

If Not conn Is Nothing Then
  conn.Close
  Set conn = Nothing
End If
%>
