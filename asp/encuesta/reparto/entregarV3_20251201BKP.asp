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
  </style> 
</head>
<body class="bg-light">

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></strong>
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
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Sel " & hojaRutaID & ", " & clienteID& ", '" & Replace(cond,"'","''") & "'"
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
if cuentaCorriente=1 Then
	totalCobrar=0.00
End If
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
	  <p><strong>Total a Cobrar:</strong> $<span id="totalGeneral"><%= FormatNumber(ImporteTotal, 2) %></span></p>
    </div>
  </div>

  <button type="button" class="btn btn-outline-secondary mb-3" onclick="restaurarCantidades()">Restablecer Cantidades</button>

 <% Do Until rs.EOF %>
    <div class="card mb-3 factura-card" data-facturaid="<%= rs("FacturaID") %>">
      <div class="card-header bg-primary text-white">
        <%= rs("Factura") %> - Total: $<span class="factura-total" data-facturaid="<%= rs("FacturaID") %>"><%= FormatNumber(rs("TotalFacturas"), 2) %></span>
		<%if rs("isOR") = 1 Then %>
		<span class="badge bg-warning text-dark ms-2">Tiene una orden de Retiro</span>
		<%End If%>
      </div>
      <ul class="list-group list-group-flush">
        <%
        facturaID = rs("FacturaID")
        sqlItems = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Detalle_Sel " & facturaID
        Set rsItems = conn.Execute(sqlItems)

        Do Until rsItems.EOF
        itemId = rsItems("Id")
        cantidad = rsItems("Cantidad")
        precioUnitario = rsItems("PrecioUnitario")
        importeConIVA = rsItems("ImporteConIVA")
        percepcionIIBB = rsItems("PercepcionIIBB")
        totalConImpuestos = rsItems("TotalConImpuestos")
        dto =  FormatNumber(rsItems("Dto"),2)
        %>
          <li class="list-group-item d-flex justify-content-between align-items-start item-factura" 
              data-itemid="<%= itemId %>" 
              data-facturaid="<%= facturaID %>" 
              data-precio="<%= FormatearParaJS(precioUnitario) %>" 
              data-iva="<%= FormatearParaJS(importeConIVA) %>" 
              data-iibb="<%= FormatearParaJS(percepcionIIBB) %>" 
              data-totalimp="<%= FormatearParaJS(totalConImpuestos) %>" 
              data-cantorig="<%= cantidad %>">
            <div class="ms-2 me-auto">
              <div class="fw-bold"><%= rsItems("Descripcion") %></div>
              <div>
				Articulo: <span class="item-articulo"><%= rsItems("Articulo") %></span> 
                Cant: <span class="item-cantidad"><%= cantidad %></span>
                <input type="hidden" name="item_<%= facturaID %>_<%= itemId %>" value="<%= cantidad %>" class="input-cantidad-item">
                <% If dto = 0 Then %>
                  <button type="button" class="btn btn-sm btn-outline-secondary ms-1 btn-restar">-</button>
                <% Else %>
                  <span class="badge bg-warning text-dark ms-2">Dto</span>
                <% End If %>
              </div>
              <small>
                Unitario: $<%= FormatNumber(precioUnitario, 2) %><br>
                IVA: $<%= FormatNumber(importeConIVA, 2) %><br>
                IIBB: $<%= FormatNumber(percepcionIIBB, 2) %><br>
                <strong>Total c/ Impuestos: $<span class="item-total-imp"><%= FormatNumber(totalConImpuestos, 2) %></span></strong>
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
<% rs.MoveNext : Loop : rs.Close %>

    <input type="hidden" name="hdrid" value="<%= hojaRutaID %>">
    <input type="hidden" name="clienteid" value="<%= clienteID %>">
    <input type="hidden" name="esCC" value="<%= cuentaCorriente %>">
    <input type="hidden" name="cond" value="<%= cond %>">

	<h4 class="mb-3">Formas de Pago</h4>
	<div class="row g-2">
	  <% 
		Dim formasPago: formasPago = Array("Efectivo", "Cheque/Echeq", "Transferencia")
		Dim icono,tooltip,statusClass
 
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
	  <% Next %>
	</div>

	<div class="mt-3 p-3 bg-white border rounded">
	  <h5>Saldo restante: $<span id="saldoRestante"><%If cuentaCorriente=0 THEN %><%= FormatNumber(totalCobrar, 2) %><%Else%> <%= FormatNumber(0.00, 2) %> <%End If%></span></h5>
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
      <button type="submit" class="btn btn-success btn-lg">Registrar Pago</button>
    </div>
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
  const facturaTotales = {};
  document.querySelectorAll('.item-factura').forEach(function (item) {
    const cantActual = parseInt(item.querySelector('.item-cantidad').textContent);
    const cantOriginal = parseInt(item.dataset.cantorig);
    const totalImpOriginal = parseFloat(item.dataset.totalimp);
    const totalItem = (totalImpOriginal / cantOriginal) * cantActual;
    item.querySelector('.item-total-imp').textContent = formatNumberAr(totalItem);
    item.querySelector('.input-cantidad-item').value = cantActual;
    if (cantActual < cantOriginal) item.classList.add('bg-warning-subtle'); else item.classList.remove('bg-warning-subtle');
    const facturaID = item.dataset.facturaid;
    if (!facturaTotales[facturaID]) facturaTotales[facturaID] = 0;
    facturaTotales[facturaID] += totalItem;
    totalGeneral += totalItem;
  });
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
  document.querySelectorAll('.item-factura').forEach(item => {
    item.querySelector('.item-cantidad').textContent = item.dataset.cantorig;
  });
  actualizarTotales();
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
document.getElementById('formPagos').addEventListener('submit',function(e){
  if(<%= cuentaCorriente %>==1){
      const dni=document.getElementById('dniCC')?.value.trim() || '';
      if(!/^\d{7,10}$/.test(dni)){
          alert('Debe ingresar un DNI válido (solo números).');
          e.preventDefault();
          return;
      }
  }else{
      const saldo=parseFloat(document.getElementById('saldoRestante').textContent.replace(/\./g,'').replace(',','.'));
      if (saldo > 0  ) {
			alert('El saldo restante no puede ser mayor a 0 para registrar el pago.');
			e.preventDefault();
		}
  }
});
document.querySelectorAll('input[type="number"]').forEach(input => {
  input.addEventListener('input', actualizarSaldo);
});
</script>

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
