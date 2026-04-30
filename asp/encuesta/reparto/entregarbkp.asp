<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<!--#include file="sidebar.asp" -->
<%
' Función para formatear números para JavaScript (sin separadores de miles, con punto decimal)
Function FormatearParaJS(ByVal valor)
  Dim s
  s = Trim(CStr(valor))
  s = Replace(s, ".", "") ' quita separador de miles (puntos)
  s = Replace(s, ",", ".") ' reemplaza coma decimal por punto
  FormatearParaJS = s
End Function

Dim hojaRutaID, clienteID, rs, sql, totalCobrar, clienteNombre, totalFacturas

hojaRutaID = Request.QueryString("hdrid")
clienteID = Request.QueryString("clienteid")

If hojaRutaID = "" Or clienteID = "" Then
  Response.Write "Faltan parámetros."
  Response.End
End If

' Obtener información del cliente y facturas
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Sel " & hojaRutaID & ", " & clienteID
Set rs = conn.Execute(sql)

If rs.EOF Then
  Response.Write "No se encontraron datos para este cliente."
  Response.End
End If

ClienteID = rs("clienteID")
clienteNombre = rs("clienteID")
totalCobrar = rs("TotalFacturas")
totalFacturas = rs("CantidadFacturas")
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Transportista</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
</head>
<body>

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>


<div class="main-content">

<h1>Entrega de Facturas</h1>

<p><strong>Hoja de Ruta:</strong> <%= hojaRutaID %></p>
<p><strong>Cliente:</strong> <%= clienteNombre %></p>
<p><strong>Facturas:</strong> <%= totalFacturas %></p>
<p><strong>Total a Cobrar:</strong> $<%= FormatNumber(totalCobrar, 2) %></p>

<%
Do Until rs.EOF
%>
  <div class="factura">
    <p><strong>Factura:</strong> <%= rs("FacturaID") %> -  
       <strong>Total:</strong> $<%= FormatNumber(rs("TotalFacturas"), 2) %></p>
    <ul class="items">
      <%
      Dim facturaID, rsItems, sqlItems
      facturaID = rs("FacturaID")
      sqlItems = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Detalle_Sel " & facturaID
      Set rsItems = conn.Execute(sqlItems)

      Do Until rsItems.EOF
      %>
        <li>
          <strong><%= rsItems("Descripcion") %></strong> - 
          Cant: <%= rsItems("Cantidad") %> - 
          Precio: $<%= FormatNumber(rsItems("PrecioUnitario"), 2) %> - 
          Total: $<%= FormatNumber(rsItems("Importe"), 2) %>
          <a class="eliminar" href="eliminar_item.asp?detalleid=<%= rsItems("Id") %>&hdrid=<%= hojaRutaID %>&clienteid=<%= clienteID %>" 
             onclick="return confirm('¿Está seguro que desea eliminar este ítem?')">
             <i class="fas fa-trash-alt"></i>
          </a>
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

<form id="formPagos" method="post" action="registrar_pago.asp">
  <input type="hidden" name="hdrid" value="<%= hojaRutaID %>">
  <input type="hidden" name="clienteid" value="<%= clienteID %>">

  <h3>Formas de Pago</h3>

  <% For i = 1 To 7 %>
    <div class="formapago">
      <label>Pago <%= i %>: $</label>
      <input id="pago<%= i %>" type="number" name="pago<%= i %>" step="0.01" min="0" value="0">
    </div>
  <% Next %>

  <div class="saldo">
    Saldo restante: $<span id="saldoRestante"><%= FormatNumber(totalCobrar, 2) %></span>
  </div>

  <br>
  <input type="submit" value="Registrar Pago">
</form>


</div>


<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
	
  function actualizarSaldo() {
    const total = <%= FormatearParaJS(totalCobrar) %>;
    let sumaPagos = 0;

    for (let i = 1; i <= 7; i++) {
      const input = document.getElementById(`pago${i}`);
      if (input) {
        const val = parseFloat(input.value.replace(",", "."));
        if (!isNaN(val)) {
          sumaPagos += val;
        }
      }
    }

    const saldoRestante = total - sumaPagos;
    document.getElementById("saldoRestante").textContent = saldoRestante.toFixed(2);
  }

  document.addEventListener("DOMContentLoaded", function () {
    for (let i = 1; i <= 7; i++) {
      const input = document.getElementById(`pago${i}`);
      if (input) {
        input.addEventListener("input", actualizarSaldo);
        input.addEventListener("change", actualizarSaldo);
      }
    }
    actualizarSaldo();

    // Opcional: evitar enviar si saldo no es 0
    document.getElementById("formPagos").addEventListener("submit", function(e) {
      const saldo = parseFloat(document.getElementById("saldoRestante").textContent.replace(",", "."));
      if (saldo.toFixed(2) != 0) {
        alert("El saldo restante debe ser 0 para registrar el pago.");
        e.preventDefault();
      }
    });
  });
</script>

</body>
</html>
