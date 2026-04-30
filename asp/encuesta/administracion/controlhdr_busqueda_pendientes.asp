<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"
%>
<!--#include file="sidebar.asp" -->

<%
Dim HDR, ClienteID, nuevaCondicion, sql, rsCliente, fecha, fechaDate
HDR = Trim(Request("hdr"))
ClienteID = Trim(Request("clienteid"))
FacturaID = Trim(Request("facturaid"))

fecha = Request("fecha")
If fecha = "" Then
    ' Fecha actual en formato yyyy-mm-dd
    fecha = Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2)


' Convertimos la fecha a tipo Date para poder operar

fechaDate = CDate(fecha)

' Restamos 7 días (una semana)
fechaDate = DateAdd("d", -7, fechaDate)

' Volvemos a convertir a string yyyy-mm-dd
fecha = Year(fechaDate) & "-" & Right("0" & Month(fechaDate), 2) & "-" & Right("0" & Day(fechaDate), 2)
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pendientes</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    <style>
        .cards-container { display: flex; flex-direction: column; gap: 20px; }
        .hdr-card {
            border: 2px solid #2980b9;
            border-radius: 10px;
            padding: 15px;
            background: #e8f4fb;
        }
        .hdr-card h2 { margin-top: 0; color: #2980b9; }
        .cliente-card {
            border: 1px solid #ccc;
            border-radius: 8px;
            background: #fff;
            padding: 15px;
            margin: 10px 0 0 20px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }
        .cliente-card h3 { margin-top: 0; }
		table.facturas {
			width: 100%;
			border-collapse: collapse;
			margin-top: 10px;
		}
		table.facturas th, table.facturas td {
			border: 1px solid #ddd;
			padding: 8px;
			text-align: center;
		}
		table.facturas th {
			background-color: #2980b9;
			color: white;
			font-weight: bold;
			padding: 10px 8px;
			text-align: center;
			border-bottom: 2px solid #1c5980;
		}
        .estado-icon { margin-right: 6px; }
        form.update-form {
            margin-top: 15px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        form.update-form label { white-space: nowrap; }
        form.update-form select { padding: 5px; }
        form.update-form button {
            cursor: pointer;
            background-color: #3498db;
            color: white;
            padding: 6px 12px;
            border: none;
            border-radius: 4px;
        }
    </style>
</head>
<body>

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout" />
    </form>
</header>

<div class="main-content">
    
	<% If Request.QueryString("msg") = "ok" Then %>
        <div style="background:#d4edda; color:#155724; padding:10px; margin-bottom:15px; border:1px solid #c3e6cb; border-radius:5px;">
            ✅ Condición de venta actualizada correctamente.
        </div>
    <% End If %>

	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		
    <h2>Busqueda de Pendientes</h2>
   
		
		<a href="/menu.asp" 
		   class="icon-btn" 
		   title="Volver al menu"
		   style="margin-left: left; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
			<i class="fas fa-arrow-left"></i> Menu
		</a>
	</div> 
	

    <form method="get" action="controlhdr_busqueda_pendientes.asp" style="margin-bottom:20px;">
        <label for="fecha">Fecha:</label>
        <input type="date" name="fecha" id="fecha" value="<%= fecha %>">

	
        <label>Hoja de Ruta:</label>
        <input type="text" name="hdr" value="<%=HDR%>" />
        &nbsp;&nbsp;
        <label>Cliente ID:</label>
        <input type="text" name="clienteid" value="<%=ClienteID%>" />
        &nbsp;&nbsp;
        <label>Factura ID:</label>
        <input type="text" name="facturaid" value="<%=FacturaID%>" />
        <input type="submit" value="Buscar" />
    </form>

<%
If Fecha<> "" or HDR <> "" Or ClienteID <> "" or FacturaID<>"" Then
    sql = "EXEC usp_Transportista_HojaDeRuta_Buscar_Pendientes '" &  Fecha  & "', "& _ 
          IIf(HDR = "", "0", HDR) & ", " & IIf(ClienteID = "", "0", ClienteID) & ", " & IIf(FacturaID = "", "0", FacturaID)& ", 0"
	'response.write sql
    Set rsCliente = conn.Execute(sql)

	If Not rsCliente.EOF Then
		Dim currentClienteID, clienteNombre, formaPagoCliente
		currentClienteID = -1

		Response.Write "<div class='cards-container'>"

		Do While Not rsCliente.EOF
			If rsCliente("ClienteID") <> currentClienteID Then
				' Cerrar tabla anterior si corresponde
				If currentClienteID <> -1 Then
					Response.Write "</tbody></table>"
					Response.Write "</div>" ' cierra cliente-card
				End If

				' Nuevo cliente
				currentClienteID = rsCliente("ClienteID")
				clienteNombre = rsCliente("ClienteNombre")
				formaPagoCliente = rsCliente("FormaPago")

				Response.Write "<div class='cliente-card'>"
				Response.Write "<h3><i class='fas fa-user'></i> " & clienteNombre & "</h3>"
				Response.Write "<p><strong>Cliente ID:</strong> " & currentClienteID & "</p>"

				Response.Write "<table class='facturas'>"
				Response.Write "<thead><tr>"
				Response.Write "<th>Hoja de Ruta</th>"
				Response.Write "<th>Fecha</th>"
				Response.Write "<th>Factura ID</th>"
				Response.Write "<th>Importe Factura</th>"
				Response.Write "<th>Forma Pago</th>"
				Response.Write "<th>Estado</th>"
				Response.Write "</tr></thead><tbody>"
			End If

			' Fila de factura
			Response.Write "<tr>"
			Response.Write "<td>" & rsCliente("HojaDeRutaID") & "</td>"
			Response.Write "<td>" & FormatDateTime(rsCliente("Fecha"),2) & "</td>"
			Response.Write "<td>" & rsCliente("FacturaID") & "</td>"
			Response.Write "<td>$" & FormatNumber(rsCliente("ImporteFactura"), 2) & "</td>"
			Response.Write "<td>" & Server.HTMLEncode(rsCliente("FormaPago")) & "</td>"
			Response.Write "<td>"

			Select Case rsCliente("Estado")
				Case 1
					Response.Write "<i class='fas fa-check-circle estado-icon' style='color:green;' title='Entregado'></i>Entregado"
				Case 2
					Response.Write "<i class='fas fa-ban estado-icon' style='color:red;' title='Anulado'></i>Anulado"
				Case 3
					Response.Write "<i class='fas fa-clock estado-icon' style='color:orange;' title='Pendiente'></i>Pendiente"
				Case Else
					Response.Write "Desconocido"
			End Select

			Response.Write "</td>"
			Response.Write "</tr>"

			rsCliente.MoveNext
		Loop

		' Cierra última tabla y tarjeta
		Response.Write "</tbody></table>"
		Response.Write "</div>" ' cliente-card
		Response.Write "</div>" ' cards-container

	Else
		Response.Write "<p style='color:red;'>❌ No se encontraron clientes.</p>"
	End If

	rsCliente.Close
	Set rsCliente = Nothing
 
End If

%>

</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>

</body>
</html>

<%
Function IfThen(cond, val)
    If cond Then
        IfThen = val
    Else
        IfThen = ""
    End If
End Function

Function IIf(condicion, valorVerdadero, valorFalso)
    If condicion Then
        IIf = valorVerdadero
    Else
        IIf = valorFalso
    End If
End Function
%>
