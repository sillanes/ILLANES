<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("NombreTransportista") = "" Then
  Response.Redirect "login.asp"
End If
%>

<!--#include file="sidebar.asp" -->
<%
Function FormatearParaJS(v)
  Dim s
  s = Trim(CStr(v))
  s = Replace(s,".","")
  s = Replace(s,",",".")
  FormatearParaJS = s
End Function

Dim hojaRutaID, clienteID, cond
hojaRutaID = Request.QueryString("hdrid")
clienteID  = Request.QueryString("clienteid")
cond       = Request.QueryString("cond")

If hojaRutaID="" Or clienteID="" Or cond="" Then
  Response.Write "Faltan parámetros."
  Response.End
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ILLANES HNOS SRL</title>

<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css">

<style>
.bg-danger-subtle{background:#f8d7da!important;}
.factura-card.border-danger{border-width:2px!important;}
.factura-total,#totalGeneral,#saldoRestante{
  min-width:120px;display:inline-block;text-align:right;
}
.formadepago-efectivo{background:#2c9c55}
.formadepago-cheque{background:#fff9e6}
.formadepago-transferencia{background:#6974f9}
.formadepago-otro{background:#1322d7}
</style>

<script>
  function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}
</script>

</head>

<body class="bg-light">

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex:1;">👤 <%=Server.HTMLEncode(Session("NombreTransportista"))%></strong>
    <form method="post" action="logout.asp" style="margin:0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="container my-4">
<form id="formPagos" method="post" action="registrar_pagoV3.asp">
<input type="hidden" name="redir_to_comprobantes" value="1">
<input type="hidden" name="return_after_comprobante" value="<%=Server.HTMLEncode("hojaderutav3.asp?hdrid=" & hojaRutaID)%>">

<%
Dim rs, sql
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Sel " & hojaRutaID & "," & clienteID & ",'" & Replace(cond,"'","''") & "'"
Set rs = conn.Execute(sql)

If rs.EOF Then
  Response.Write "No se encontraron datos."
  Response.End
End If

Dim clienteNombre, totalFacturas, ImporteTotal, totalCobrar, cuentaCorriente
clienteNombre   = rs("clienteNombre")
totalFacturas   = rs("CantidadFacturas")
ImporteTotal    = rs("ImporteTotal")
cuentaCorriente = rs("CuentaCorriente")
totalCobrar     = ImporteTotal
If cuentaCorriente = 1 Then totalCobrar = 0
%>

<div class="d-flex justify-content-between align-items-center mb-4">
  <h2>Pedidos</h2>
  <a href="hojaderutav3.asp?hdrid=<%= hojaRutaID %>"
     class="icon-btn"
     style="background:#2e30c7;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none">
    <i class="fas fa-arrow-left"></i> Volver
  </a>
</div>

<% If cuentaCorriente = 1 Then %>
<div class="alert alert-warning">
  ⚠️ Cliente en <b>Cuenta Corriente</b>.
  Las facturas o productos descontados quedarán pendientes.
</div>

<div class="mb-3">
  <label class="form-label">DNI de quien recibe</label>
  <input type="text" inputmode="numeric" pattern="[0-9]{7,10}"
         class="form-control" id="dniCC" name="dni" required>
</div>
<% End If %>

<h4 class="mb-3">Formas de Pago</h4>

<!-- ✅ NOTA COMO PEDISTE -->
<div class="alert alert-info py-2 mb-3" style="font-size:0.95rem">
  ℹ️ <strong>Usar</strong> <code>,</code> como decimal
  <br>
  <small class="text-muted">Ejemplo: <strong><%=replace(FormatNumber(ImporteTotal,2),".","")%></strong></small>
</div>

<div class="row g-2">
<%
Dim formasPago: formasPago = Array("Efectivo","Cheque/Echeq","Transferencia")

Dim i, statusClass
For i = 0 To UBound(formasPago)

  Select Case i
    Case 0:statusClass="formadepago-efectivo"
    Case 1:statusClass="formadepago-cheque"
    Case 2:statusClass="formadepago-transferencia"
    Case Else: statusClass=""
  End Select
%>

<div class="col-6 col-md-4 <%=statusClass%>">
  <label class="form-label"><%=formasPago(i)%></label>
  <input id="pago<%=i+1%>" name="pago<%=i+1%>"
         type="text"
         inputmode="decimal"
         pattern="[0-9]*[,.]?[0-9]*"
         class="form-control"
         value="0">
  <input type="hidden" name="codigo<%=i+1%>" value="<%=i+1%>">
</div>
<% Next %>
</div>

<div class="mt-3 p-3 bg-white border rounded">
<b>Saldo restante:</b> $
<span id="saldoRestante"><%=FormatNumber(totalCobrar,2)%></span>
</div>

<div class="d-grid mt-4">
<button type="button" id="btnRegistrar" class="btn btn-success btn-lg">
Registrar Pago
</button>
</div>

<div class="card mb-3 mt-4">
<div class="card-body">
  <p><b>Hoja de Ruta:</b> <%=hojaRutaID%></p>
  <p><b>Cliente:</b> <%=clienteID%> - <%=clienteNombre%></p>
  <p><b>Facturas:</b> <%=totalFacturas%></p>
  <p><b>Total:</b> $
    <span id="totalGeneral" data-totalorig="<%=FormatearParaJS(ImporteTotal)%>">
      <%=FormatNumber(ImporteTotal,2)%>
    </span>
  </p>
</div>
</div>

<button type="button" class="btn btn-outline-secondary mb-3"
        onclick="restaurarCantidades()">Restablecer Cantidades</button>

<%
Do Until rs.EOF
  Dim facturaID, esNC
  facturaID = rs("FacturaID")
  esNC = (Left(rs("Factura"),2)="NC")
%>

<div class="card mb-3 factura-card"
     data-facturaid="<%=facturaID%>"
     data-totalorig="<%=FormatearParaJS(rs("TotalFacturas"))%>">

<div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">
  <div>
    <%=rs("Factura")%> –
    Total: $
    <span class="factura-total" data-facturaid="<%=facturaID%>">
      <%=FormatNumber(rs("TotalFacturas"),2)%>
    </span>
    <% If esNC Then %>
      <span class="badge bg-info ms-2">Nota de Crédito</span>
    <% End If %>
  </div>

  <div class="d-flex gap-2">
    <% If Not esNC Then %>
    <button type="button" class="btn btn-sm btn-outline-light"
            onclick="toggleItems('<%=facturaID%>')">
      <i class="fas fa-list"></i>
    </button>
    <% End If %>

    <button type="button" class="btn btn-sm btn-danger"
            onclick="abrirModalAnular('<%=facturaID%>')">
      <i class="fas fa-times"></i>
    </button>
  </div>
</div>

<ul class="list-group list-group-flush factura-items"
    id="items_<%=facturaID%>"
    data-loaded="0"
    style="display:none"></ul>

</div>

<%
rs.MoveNext
Loop
rs.Close
%>

<!-- HIDDEN -->
<input type="hidden" name="hdrid" value="<%=hojaRutaID%>">
<input type="hidden" name="clienteid" value="<%=clienteID%>">
<input type="hidden" name="esCC" value="<%=cuentaCorriente%>">
<input type="hidden" name="cond" value="<%=cond%>">

<!-- GEO -->
<input type="hidden" name="latitud"  id="latitud">
<input type="hidden" name="longitud" id="longitud">
<input type="hidden" id="geo_ok" value="0">

<!-- MODAL ANULAR -->
<div id="modalAnular" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483647">
  <div style="max-width:420px;margin:20vh auto;background:#fff;border-radius:10px">
    <div style="padding:14px;font-weight:700">Confirmar anulación</div>
    <div style="padding:14px">¿Anular esta factura completa?</div>
    <div style="display:flex;gap:10px;justify-content:flex-end;padding:12px">
      <button type="button" class="btn btn-outline-secondary" onclick="cerrarModalAnular()">Cancelar</button>
      <button type="button" class="btn btn-danger" onclick="confirmarAnulacion()">Anular</button>
    </div>
  </div>
</div>

</form>
</div>

<script>
function formatNumberAr(n){
  n = Number(n) || 0;
  return n.toLocaleString('es-AR',{minimumFractionDigits:2, maximumFractionDigits:2});
}

/**
 * Parser robusto:
 * - Si vienen '.' y ',' juntos, toma como decimal el último que aparece.
 *   Ej: 101.103,53  -> decimal ','  (AR)
 *   Ej: 101,103.53  -> decimal '.'  (US)
 * - Si viene solo ',' -> decimal ','
 * - Si viene solo '.' -> decimal '.'
 */
function parseMoney(str){
  if(!str) return 0;
  str = String(str).trim();
  if(str === "") return 0;

  // quitar $ y espacios
  str = str.replace(/\$/g,"").replace(/\s+/g,"");

  const hasDot = str.indexOf(".") !== -1;
  const hasComma = str.indexOf(",") !== -1;

  if(hasDot && hasComma){
    const lastDot = str.lastIndexOf(".");
    const lastComma = str.lastIndexOf(",");
    const decSep = (lastDot > lastComma) ? "." : ",";
    const thouSep = (decSep === ".") ? "," : ".";

    // sacar miles
    str = str.split(thouSep).join("");
    // pasar decimal a punto
    if(decSep === ",") str = str.replace(",", ".");
    return parseFloat(str) || 0;
  }

  if(hasComma && !hasDot){
    // Asumimos coma decimal
    str = str.replace(/\./g,""); // por si meten puntos como miles
    str = str.replace(",", ".");
    return parseFloat(str) || 0;
  }

  // solo punto (decimal) o número limpio
  str = str.replace(/,/g,""); // por si meten coma como miles
  return parseFloat(str) || 0;
}

/* ===========================================================
   ✅ NUEVO: Validación "solo registrar si saldo restante es 0"
   - Respeta Cuenta Corriente (saldo fijo 0)
   - Usa tolerancia 0.01 por redondeos
=========================================================== */
function saldoEsCero(){
  if(<%= cuentaCorriente %> == 1) return true; // CC: siempre permite

  const saldoTxt = document.getElementById('saldoRestante').textContent;
  const saldo = parseMoney(saldoTxt);
  return Math.abs(saldo) < 0.01;
}

/* ================= TOTALES ================= */
function actualizarTotales(){
  let totalGeneral=0;

  document.querySelectorAll('.factura-card').forEach(card=>{
    let totalFactura=0;

    if(card.classList.contains('border-danger')){
      totalFactura=0;
    }else{
      const items=card.querySelectorAll('.item-factura');
      if(items.length){
        items.forEach(it=>{
          const cant = parseInt(it.querySelector('.item-cantidad').textContent,10) || 0;
          const orig = parseInt(it.dataset.cantorig,10) || 1;
          const tot  = parseFloat(it.dataset.totalimp) || 0;

          const val = (tot/orig)*cant;
          const totalEl = it.querySelector('.item-total-imp');
          if(totalEl) totalEl.textContent = formatNumberAr(val);
          totalFactura += val;
        });
      }else{
        totalFactura = parseFloat(card.dataset.totalorig || 0) || 0;
      }
    }

    const ft = card.querySelector('.factura-total');
    if(ft) ft.textContent = formatNumberAr(totalFactura);
    totalGeneral += totalFactura;
  });

  document.getElementById('totalGeneral').textContent = formatNumberAr(totalGeneral);

  // ✅ SIEMPRE recalcular saldo (pero CC lo fija en 0)
  actualizarSaldo();
}

function actualizarSaldo(){
  // 👉 CUENTA CORRIENTE: saldo fijo en 0
  if(<%= cuentaCorriente %> == 1){
    document.getElementById('saldoRestante').textContent = formatNumberAr(0);

    // ✅ (opcional recomendado) habilitar botón
    const btn = document.getElementById("btnRegistrar");
    if(btn) btn.disabled = false;

    return;
  }

  const total = parseMoney(document.getElementById('totalGeneral').textContent);

  let pagos = 0;
  for(let i=1;i<=3;i++){
    pagos += parseMoney(document.getElementById('pago'+i).value);
  }

  let saldo = Math.round((total - pagos) * 100) / 100;
  document.getElementById('saldoRestante').textContent = formatNumberAr(saldo);

  // ✅ (opcional recomendado) deshabilitar botón si saldo != 0
  const btn = document.getElementById("btnRegistrar");
  if(btn) btn.disabled = !saldoEsCero();
}

function restaurarCantidades(){
  document.querySelectorAll('.item-factura').forEach(i=>{
    const orig = i.dataset.cantorig;
    const sp = i.querySelector('.item-cantidad');
    const inp = i.querySelector('.input-cantidad-item');
    if(sp) sp.textContent = orig;
    if(inp) inp.value = orig;
  });
  document.querySelectorAll('.factura-card').forEach(c=>{
    c.classList.remove('border','border-danger');
  });
  actualizarTotales();
}

/* ================= LAZY ITEMS ================= */
function toggleItems(fid){
  const ul=document.getElementById("items_"+fid);
  if(!ul) return;

  // marcar factura abierta
  if(!document.querySelector('input[name="factura_abierta_'+fid+'"]')){
    const h=document.createElement("input");
    h.type="hidden";
    h.name="factura_abierta_"+fid;
    h.value="1";
    document.getElementById("formPagos").appendChild(h);
  }

  ul.style.display = (ul.style.display==="none" ? "block" : "none");
  if(ul.dataset.loaded==="1") return;

  ul.innerHTML="<li class='list-group-item text-center'><i class='fas fa-spinner fa-spin'></i> Cargando…</li>";

  fetch("factura_items_ajax.asp?facturaid="+fid)
    .then(r=>r.text())
    .then(html=>{
      ul.innerHTML=html;
      ul.dataset.loaded="1";

      ul.querySelectorAll(".btn-restar").forEach(btn=>{
        btn.addEventListener("click",()=>{
          const li = btn.closest(".item-factura");
          if(!li) return;

          const span = li.querySelector(".item-cantidad");
          const input = li.querySelector(".input-cantidad-item");
          let val = parseInt(span ? span.textContent : "0",10) || 0;
          val = Math.max(0,val-1);
          if(span) span.textContent = val;
          if(input) input.value = val;

          actualizarTotales();
        });
      });

      actualizarTotales();
    });
}

/* ================= GEO + SUBMIT ================= */
document.getElementById("btnRegistrar").addEventListener("click",function(){

  // ✅ NUEVO: NO registrar si el saldo restante no es 0
  if(!saldoEsCero()){
    alert("⚠️ No se puede registrar el pago.\nEl saldo restante debe ser 0,00.");
    return;
  }

  if(<%=cuentaCorriente%>===1){
    const dni=document.getElementById("dniCC");
    if(!dni||dni.value.trim()===""){
      alert("Debe ingresar el DNI.");
      dni.focus();
      return;
    }
  }

  if(document.getElementById("geo_ok").value==="1"){
    document.getElementById("formPagos").submit();
    return;
  }

  navigator.geolocation.getCurrentPosition(
    pos=>{
      document.getElementById("latitud").value  = pos.coords.latitude.toFixed(6);
      document.getElementById("longitud").value = pos.coords.longitude.toFixed(6);
      document.getElementById("geo_ok").value   = "1";
      document.getElementById("formPagos").submit();
    },
    ()=>{
      alert("⚠️ No se pudo obtener la ubicación.\nAbrí el link en Chrome.");
    },
    {enableHighAccuracy:true,timeout:15000,maximumAge:0}
  );
});

// recalcular saldo al escribir pagos
["pago1","pago2","pago3"].forEach(id=>{
  const el=document.getElementById(id);
  if(el){
    el.addEventListener("input",actualizarSaldo);
    el.addEventListener("change",actualizarSaldo);
  }
});

// ✅ inicial
actualizarSaldo();
</script>

<script>
var facturaPendiente=null;

function abrirModalAnular(fid){
  facturaPendiente=fid;
  document.getElementById("modalAnular").style.display="block";
}
function cerrarModalAnular(){
  facturaPendiente=null;
  document.getElementById("modalAnular").style.display="none";
}
function confirmarAnulacion(){
  if(!facturaPendiente) return;

  if(!document.querySelector('input[name="anulada_'+facturaPendiente+'"]')){
    const h=document.createElement("input");
    h.type="hidden";
    h.name="anulada_"+facturaPendiente;
    h.value="1";
    document.getElementById("formPagos").appendChild(h);
  }

  document.querySelectorAll('.item-factura[data-facturaid="'+facturaPendiente+'"]')
    .forEach(i=>{
      const c=i.querySelector('.item-cantidad');
      const inp=i.querySelector('.input-cantidad-item');
      const t=i.querySelector('.item-total-imp');
      if(c) c.textContent="0";
      if(inp) inp.value="0";
      if(t) t.textContent="0,00";
    });

  const card=document.querySelector('.factura-card[data-facturaid="'+facturaPendiente+'"]');
  if(card) card.classList.add("border","border-danger");

  actualizarTotales();
  cerrarModalAnular();
}
</script>

</body>
</html>

<%
If Not conn Is Nothing Then conn.Close
Set conn = Nothing
%>