<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"
%>
<!--#include file="sidebar.asp" -->

<%
Dim HDR, ClienteID, nuevaCondicion, sql, rsCliente
HDR = Trim(Request("hdr"))
ClienteID = Trim(Request("clienteid"))
FacturaID = Trim(Request("facturaid"))


%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Terminadas</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
   <style>
        .resumen-totales {
            margin: 20px 0;
            padding: 15px;
            background-color: #f1f1f1;
            border-radius: 8px;
        }
        .detalle-toggle {
            cursor: pointer;
            color: #007bff;
            text-decoration: underline;
        }
        .detalle-content {
            display: none;
            margin-top: 5px;
            white-space: pre-wrap;
            font-size: 0.9em;
        }
        .subtotal {
            font-weight: bold;
            background-color: #f1f1f1;
        }
        .action-buttons {
            display: flex;
            gap: 5px;
        }
        .btn-action {
            background-color: #e0e0e0;
            border: none;
            padding: 4px 8px;
            border-radius: 4px;
            cursor: pointer;
        }
        .export-btn {
            background-color: #28a745;
            color: #fff;
            padding: 8px 12px;
            border: none;
            border-radius: 4px;
            text-decoration: none;
            margin-bottom: 15px;
            display: inline-block;
        }
        /* 脥conos tama帽o y hover */
        .action-buttons a {
            font-size: 1.3em;
            transition: color 0.2s;
        }
        .action-buttons a:hover {
            opacity: 0.7;
        }
        /* Texto validado con 铆cono */
        .validado-span {
            color: green;
            font-weight: bold;
            display: flex;
            align-items: center;
            gap: 5px;
        }
    </style>
    <script>
        function toggleDetalle(id) {
            var div = document.getElementById(id);
            div.style.display = div.style.display === 'none' || div.style.display === '' ? 'block' : 'none';
        }
        function toggleSidebar() {
            document.querySelector('.sidebar').classList.toggle('open');
        }
    </script>
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
   	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		
    <h2>Busqueda de Terminadas</h2>
		<a href="/menu.asp" 
		   class="icon-btn" 
		   title="Cerrar hoja de ruta"
		   style="margin-left: left; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
			<i class="fas fa-arrow-left"></i> Menu
		</a>
	</div> 
<!--
    <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;"> 
        <a href="exportar_excel_hdr.asp?hdrid=<%= hdrid %>" 
           class="icon-btn"   
           style="background-color: #28a745; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none; margin-left: 10px;">
           <i class="fa-solid fa-file-csv"></i>Exportar a Excel
        </a>
 
    </div> 
-->	
    <form method="get" action="controlhdr_busqueda_terminadas.asp" style="margin-bottom:20px;">
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
 
	<% If clng("0" & HDR) <>0  or ClienteID<>"" or FacturaID<>"" Then%>
		  
    <table class="table table-striped">
        <tbody>
            <% 
			sql = "EXEC usp_Transportista_HojaDeRuta_Buscar_Cerradas " & _
			IIf(HDR = "", "0", HDR) & ", " & IIf(ClienteID = "", "0", ClienteID) & ", " & IIf(FacturaID = "", "0", FacturaID)
			' response.write sql
			Set rs = conn.Execute(sql)
           Dim totalCobrar, totalCobrado
            Dim subtotalCobrar, subtotalCobrado
            Dim subtotalEfectivo, subtotalCheque, subtotalTransferencia, subtotalDescuentos, subtotalDiferencia
            Dim totalEfectivo, totalCheque, totalTransferencia, totalDescuentos, totalDiferencia, estadoentrega
            totalCobrar = 0.00
            totalCobrado = 0.00
            subtotalCobrar = 0.00
            subtotalCobrado = 0.00
            subtotalEfectivo = 0.00
            subtotalCheque = 0.00
            subtotalTransferencia = 0.00
            subtotalDescuentos = 0.00
            subtotalDiferencia = 0.00
            totalEfectivo = 0.00
            totalCheque = 0.00
            totalTransferencia = 0.00
            totalDescuentos = 0.00
            totalDiferencia = 0.00
 
            Dim currentForma, nuevaForma
            Dim diferencia, detalleTexto, detalleID
            currentForma = ""
            Do Until rs.EOF
                nuevaForma = rs("FormaPago")
                If nuevaForma <> currentForma Then
                    If currentForma <> "" Then
                        Response.Write "<tr class='subtotal'>"
                        Response.Write "<td colspan='4'>Subtotal</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalCobrar, 2) & "</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalCobrado, 2) & "</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalEfectivo, 2) & "</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalCheque, 2) & "</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalTransferencia, 2) & "</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalDescuentos, 2) & "</td>"
                        Response.Write "<td>$" & FormatNumber(subtotalDiferencia, 2) & "</td>"
                        Response.Write "<td> </td>" 
                        Response.Write "<td> </td>" 
                        Response.Write "<td></td><td></td></tr></tbody></table><br>"
                    End If
            %>
                <h2><%= nuevaForma %></h2>
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th rowspan="2">HojaDeRutaID</th>
                            <th rowspan="2">Cliente</th>
                            <th rowspan="2">Facturas</th>
							<th rowspan="2">Comp</th>
                            <th rowspan="2">Total a Cobrar</th>
                            <th colspan="4">Total Cobrado</th>
                            <th rowspan="2">Descuentos</th>
                            <th rowspan="2">Diferencia</th>
                            <th rowspan="2">Observacion</th>
                            <!-- <th rowspan="2">Modulo Transferencias</th> -->
                            <th rowspan="2">Acciones</th>
                            <th rowspan="2">Facturas</th>
                            <th rowspan="2">Validar</th>
                        </tr>
                        <tr>
                            <th>Total</th>
                            <th>Efectivo</th>
                            <th>Cheque</th>
                            <th>Transferencia</th>
                        </tr>            
                    </thead>
                    <tbody>
            <%
                    currentForma = nuevaForma
                    subtotalCobrar = 0
                    subtotalCobrado = 0
                    subtotalEfectivo = 0
                    subtotalCheque = 0
                    subtotalTransferencia = 0
                    subtotalDescuentos = 0
                    subtotalDiferencia = 0
                End If

                Dim impCobrar, impCobrado, descuentos, efectivo, cheque, transferencia
                impCobrar = CDbl(rs("ImporteACobrar"))
                impCobrado = CDbl(rs("TotalCobrado"))
                descuentos = CDbl(rs("Descuentos"))
                efectivo = CDbl(rs("Efectivo"))
                cheque = CDbl(rs("Cheque"))
                transferencia = CDbl(rs("Transferencia"))
                diferencia = impCobrar - impCobrado - descuentos
                detalleID = "detalle_" & rs("ClienteID")
                detalleTexto = Trim(rs("Detalles"))
                Facturas = Trim(rs("facturas"))
                FacturaID = "Facturas_" & rs("ClienteID")& Facturas

                totalCobrar = totalCobrar +  impCobrar 
                totalCobrado = totalCobrado + efectivo + cheque + transferencia
                subtotalCobrar = subtotalCobrar + impCobrar
                subtotalCobrado = subtotalCobrado + efectivo + cheque + transferencia
                totalEfectivo = totalEfectivo + efectivo
                totalCheque = totalCheque + cheque
                totalTransferencia = totalTransferencia + transferencia
                totalDescuentos = totalDescuentos + descuentos
                totalDiferencia = totalDiferencia + IIF(rs("estado")>=2,0,diferencia)
                subtotalEfectivo = subtotalEfectivo + efectivo
                subtotalCheque = subtotalCheque + cheque
                subtotalTransferencia = subtotalTransferencia + transferencia
                subtotalDescuentos = subtotalDescuentos + descuentos
                subtotalDiferencia = subtotalDiferencia + IIF(rs("estado")>=2,0,diferencia)
            %>
                    <tr>
                        <td style="background-color: <%= rs("color") %>"><%= rs("HojaDeRutaID") %></td>
                        <td style="background-color: <%= rs("color") %>"><%= rs("ClienteID") & " - " & rs("ClienteNombre") %></td>
                        <td style="background-color: <%= rs("color") %>"><%= rs("TotalFacturas") %></td>
						<td style="background-color: <%= rs("color") %>; text-align:center;">
							<% If rs("CantidadComprobantes") > 0 Then %>
								<a href="comprobantes_cliente.asp?hdrid=<%= HDR %>&clienteid=<%= rs("ClienteID") %>" 
								   title="Ver comprobantes">
									<i class="fas fa-paperclip"></i>
									<strong><%= rs("CantidadComprobantes") %></strong>
								</a>
							<% Else %>
								<span style="color:#999;">0</span>
							<% End If %>
						</td>						
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(impCobrar, 2) %></td>
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(efectivo + cheque + transferencia, 2) %></td>
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(efectivo, 2) %></td>
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(cheque, 2) %></td>
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(transferencia, 2) %></td>
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(descuentos, 2) %></td>
                        <td style="background-color: <%= rs("color") %>">$<%= FormatNumber(diferencia, 2) %></td>
                        <td style="background-color: <%= rs("color") %>"><%= rs("EstadoEntrega") &"  "&  rs("Observacion") %></td>
                        <!-- Columna Acciones -->
                        <td style="background-color: <%= rs("color") %>">
                            <div class="action-buttons">
                                <% If detalleTexto <> "" Then %>
                                    <span class="detalle-toggle" onclick="toggleDetalle('<%= detalleID %>')">Ver detalles</span> 
                                <% End If %>
                            </div>
                            <% If detalleTexto <> "" Then %>
                                <div id="<%= detalleID %>" class="detalle-content"><%= Replace(detalleTexto, " - ", vbCrLf) %></div>
                            <% End If %>
                        </td>
						<td style="background-color: <%= rs("color") %>"> 
                            <div class="action-buttons">
                                <% If Facturas <> "" Then %>
                                    <span class="detalle-toggle" onclick="toggleDetalle('<%= FacturaID  %>')">Facturas</span> 
                                <% End If %>
                            </div>
                            <% If Facturas <> "" Then %>
                                <div id="<%= FacturaID %>" class="detalle-content"><%= Replace(Facturas, " - ", vbCrLf) %></div>
                            <% End If %>						
						</td>
                        <!-- Nueva columna Validaci贸n sin fondo -->
<td>
    <div class="action-buttons">
        <% 
            Dim valEstado
            valEstado = rs("Validado")

            If valEstado = 1 Then
                ' Mostrar como Validado + bot贸n de editar
        %>
                 
				<span title="<%= Server.HTMLEncode(rs("ValidadoMensaje")) %>" class="validado-span">
                    <i class="fas fa-check-circle" style="color: green;"></i> Revisado
                </span>

        <% 
            ElseIf valEstado = 0 Then
                ' Mostrar como No Validado + bot贸n de editar
        %>
 
                <span title="<%= Server.HTMLEncode(rs("ValidadoMensaje")) %>" class="validado-span" style="color: red;">
                    <i class="fas fa-times-circle"></i> Revisado
                </span>
 
        <% 
            End If 
        %>
    </div>
</td>
                    </tr>
            <%
                rs.MoveNext
            Loop
            If currentForma <> "" Then
                Response.Write "<tr class='subtotal'>"
                Response.Write "<td colspan='4'>Subtotal</td>"
                Response.Write "<td>$" & FormatNumber(subtotalCobrar, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalCobrado, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalEfectivo, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalCheque, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalTransferencia, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalDescuentos, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalDiferencia, 2) & "</td>"
                Response.Write "<td></td><td></td><td></td></tr></tbody></table><br>"
            End If
            rs.Close
            Set rs = Nothing
            %>

            <div class="resumen-totales">
                <strong>Total a Cobrar:</strong> $<%= FormatNumber(totalCobrar, 2) %><br>
                <strong>Total Cobrado:</strong> $<%= FormatNumber(totalCobrado, 2) %><br>
                <strong>Efectivo:</strong> $<%= FormatNumber(totalEfectivo, 2) %><br>
                <strong>Cheque:</strong> $<%= FormatNumber(totalCheque, 2) %><br>
                <strong>Transferencia:</strong> $<%= FormatNumber(totalTransferencia, 2) %><br>
                <strong>Descuentos:</strong> $<%= FormatNumber(totalDescuentos, 2) %><br> 
            </div> 
		<%End If%>
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
