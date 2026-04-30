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

hojaRutaID = Request.QueryString("hdrid")
clienteID  = Request.QueryString("clienteid")
cond       = Request.QueryString("cond")

If hojaRutaID = "" Or clienteID = "" Or cond = "" Then
  Response.Write "Faltan par?tros."
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
.bg-danger-subtle { background:#f8d7da!important; }
.factura-card.border-danger { border-width:2px!important; }
.factura-total,#totalGeneral,#saldoRestante{
  min-width:110px; display:inline-block; text-align:right;
}
</style>
</head>

<body class="bg-light">

<header>
  <button class="menu-toggle" onclick="toggleSidebar()">
    <i class="fas fa-bars"></i>
  </button>

  <strong style="flex:1;">
    ?? <%=Server.HTMLEncode(Session("NombreTransportista"))%>
  </strong>

  <form method="post" action="logout.asp" style="margin:0">
    <input type="submit" value="Cerrar sesi??class="logout">
  </form>
</header>

<div class="container my-4">
<form id="formPagos" method="post" action="registrar_pagoV3.asp">

<%
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Sel " & hojaRutaID & "," & clienteID & ",'" & Replace(cond,"'","''") & "'"
Set rs = conn.Execute(sql)

If rs.EOF Then
  Response.Write "No se encontraron datos."
  Response.End
End If

clienteNombre   = rs("clienteNombre")
totalCobrar     = rs("ImporteTotal")
ImporteTotal    = rs("ImporteTotal")
totalFacturas   = rs("CantidadFacturas")
cuentaCorriente = rs("CuentaCorriente")
If cuentaCorriente = 1 Then totalCobrar = 0
%>

<div class="card mb-3">
<div class="card-body">
  <p><b>Hoja de Ruta:</b> <%=hojaRutaID%></p>
  <p><b>Cliente:</b> <%=clienteID%> - <%=clienteNombre%></p>
  <p><b>Facturas:</b> <%=totalFacturas%></p>
  <p><b>Total a Cobrar:</b> $ <span id="totalGeneral" data-totalorig="<%=FormatearParaJS(ImporteTotal)%>">
      <%=FormatNumber(ImporteTotal,2)%>
    </span>
  </p>
</div>
</div>

<button type="button" class="btn btn-outline-secondary mb-3" onclick="restaurarCantidades()">
  Restablecer Cantidades
</button>

<%
Do Until rs.EOF
  facturaID = rs("FacturaID")
  esNC = (Left(rs("Factura"),2)="NC")
%>

<div class="card mb-3 factura-card"
     data-facturaid="<%=facturaID%>"
     data-totalorig="<%=FormatearParaJS(rs("TotalFacturas"))%>">

<div class="card-header bg-primary text-white d-flex justify-content-between align-items-center">

  <div>
    <%=rs("Factura")%> -
    Total: $
    <span class="factura-total" data-facturaid="<%=facturaID%>">
      <%=FormatNumber(rs("TotalFacturas"),2)%>
    </span>

    <% If esNC Then %>
      <span class="badge bg-info ms-2">Nota de Cr?to</span>
    <% End If %>
  </div>

  <div class="d-flex gap-2">

    <% If Not esNC Then %>
    <!-- VER ?EMS -->
    <button type="button"
            class="btn btn-sm btn-outline-light"
            title="Ver ?ms"
            onclick="toggleItems('<%=facturaID%>')">
      <i class="fas fa-list"></i>
    </button>
    <% End If %>

    <!-- ANULAR FACTURA -->
    <button type="button"
            class="btn btn-sm btn-danger"
            title="Anular factura"
            onclick="abrirModalAnular('<%=facturaID%>')">
      <i class="fas fa-times"></i>
    </button>

  </div>
</div>

<ul class="list-group list-group-flush factura-items"
    id="items_<%=facturaID%>"
    data-loaded="0"
    style="display:none;"></ul>

</div>

<%
rs.MoveNext
Loop
rs.Close
%>

<input type="hidden" name="hdrid" value="<%=hojaRutaID%>">
<input type="hidden" name="clienteid" value="<%=clienteID%>">
<input type="hidden" name="esCC" value="<%=cuentaCorriente%>">
<input type="hidden" name="cond" value="<%=cond%>">

<h4 class="mt-4">Formas de Pago</h4>

<div class="row g-2">
<%
formas = Array("Efectivo","Cheque/Echeq","Transferencia")
For i=0 To UBound(formas)
%>
<div class="col-6 col-md-4">
  <label class="form-label"><%=formas(i)%></label>
  <input type="text" inputmode="numeric" name="pago<%=i+1%>"
         class="form-control" value="0">
</div>
<% Next %>
</div>

<div class="mt-3 p-3 bg-white border rounded">
<b>Saldo restante:</b> $
<span id="saldoRestante"><%=FormatNumber(totalCobrar,2)%></span>
</div>

<div class="d-grid mt-4">
<button type="submit" class="btn btn-success btn-lg">
Registrar Pago
</button>
</div>

</form>
</div>

<script>
function parseFloatLocal(v){
  return parseFloat(v.replace(/\./g,'').replace(',','.'))||0;
}
function formatNumberAr(n){
  return n.toLocaleString('es-AR',{minimumFractionDigits:2});
}

function actualizarTotales(){
  let totalGeneral=0;
  document.querySelectorAll('.factura-card').forEach(card=>{
    let totalFactura=0;
    const items=card.querySelectorAll('.item-factura');
    if(items.length){
      items.forEach(item=>{
        const cant=parseInt(item.querySelector('.item-cantidad').textContent);
        const orig=parseInt(item.dataset.cantorig);
        const tot=parseFloat(item.dataset.totalimp);
        const val=(tot/orig)*cant;
        item.querySelector('.item-total-imp').textContent=formatNumberAr(val);
        totalFactura+=val;
      });
    } else {
      totalFactura=parseFloat(card.dataset.totalorig||0);
    }
    card.querySelector('.factura-total').textContent=formatNumberAr(totalFactura);
    totalGeneral+=totalFactura;
  });
  document.getElementById('totalGeneral').textContent=formatNumberAr(totalGeneral);
  actualizarSaldo();
}

function actualizarSaldo(){
  let total=parseFloatLocal(document.getElementById('totalGeneral').textContent);
  let pagos=0;
  document.querySelectorAll('[name^=pago]').forEach(i=>{
    pagos+=parseFloatLocal(i.value);
  });
  document.getElementById('saldoRestante').textContent=formatNumberAr(total-pagos);
}

function restaurarCantidades(){
  document.querySelectorAll('.item-factura').forEach(i=>{
    i.querySelector('.item-cantidad').textContent=i.dataset.cantorig;
  });
  actualizarTotales();
}

function toggleItems(fid){
  const ul=document.getElementById("items_"+fid);
  if(!ul)return;
  ul.style.display=ul.style.display==="none"?"block":"none";
  if(ul.dataset.loaded==="1")return;

  ul.innerHTML="<li class='list-group-item text-center'><i class='fas fa-spinner fa-spin'></i> Cargando...</li>";

  fetch("factura_items_ajax.asp?facturaid="+fid)
    .then(r=>r.text())
    .then(h=>{
      ul.innerHTML=h;
      ul.dataset.loaded="1";
      ul.querySelectorAll(".btn-restar").forEach(b=>{
        b.onclick=function(){
          const s=this.closest("li").querySelector(".item-cantidad");
          s.textContent=Math.max(0,parseInt(s.textContent)-1);
          actualizarTotales();
        }
      });
      actualizarTotales();
    });
}
</script>

<!-- ===== MODAL ANULAR ===== -->
<div id="modalAnular"
     style="display:none; position:fixed; inset:0;
            background:rgba(0,0,0,.55); z-index:2147483647;">

  <div style="max-width:420px; margin:20vh auto; background:#fff;
              border-radius:10px; overflow:hidden;">
    <div style="padding:14px 16px; font-weight:700;">Confirmar anulaci??div>
    <div style="padding:14px 16px;">
      ¿Seguro que querés <b>anular esta factura completa</b>?
    </div>
    <div style="display:flex; gap:10px; justify-content:flex-end; padding:12px 16px;">
      <button type="button" class="btn btn-outline-secondary"
              onclick="cerrarModalAnular()">Cancelar</button>
      <button type="button" class="btn btn-danger"
              onclick="confirmarAnulacion()">Anular</button>
    </div>
  </div>
</div>

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
  if(!facturaPendiente)return;

  if(!document.querySelector('input[name="anulada_'+facturaPendiente+'"]')){
    var h=document.createElement("input");
    h.type="hidden";
    h.name="anulada_"+facturaPendiente;
    h.value="1";
    document.getElementById("formPagos").appendChild(h);
  }

  document.querySelectorAll('.item-factura[data-facturaid="'+facturaPendiente+'"]')
    .forEach(i=>{
      i.querySelector('.item-cantidad').textContent="0";
      i.querySelector('.item-total-imp').textContent="0,00";
    });

  var card=document.querySelector('.factura-card[data-facturaid="'+facturaPendiente+'"]');
  if(card)card.classList.add("border","border-danger");

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
