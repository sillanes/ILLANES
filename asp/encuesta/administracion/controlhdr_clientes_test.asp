<!--#include file="conexion.asp" -->

<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

If Request.QueryString.Count>0 Then
    Set objRequest = Request.QueryString
Else
    Set objRequest = Request.Form
End If

Dim hdrid, fecha, transportista
hdrid = 85564 
fecha = objRequest("fecha")
mensajeOk = objRequest("msgAutoSave")
transportista = objRequest("transportista") 
transportista = clng(transportista)
If transportista = "" Then
    transportista = 0
End If

If hdrid = "" Then
    Response.Redirect "controlhdrv.asp"
End If

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

Dim nombretransportistas
Set rs = conn.Execute("EXEC [dbo].[usp_HojaDeRuta_Transportista_sel] " & hdrid)
If Not rs.EOF Then
    nombretransportistas = rs("nombre")
End If
rs.Close
Set rs = Nothing
%>
<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Clientes de Hoja de Ruta</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    <style>
        .message-ok {
            background-color: #d4edda;
            color: #155724;
            padding: 12px;
            border-radius: 8px;
            margin-top: 20px;
        }	
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
		.cerrar-btn {
			background-color: #ff030e;
			color: white;
			padding: 8px 12px;
			border-radius: 4px;
			text-decoration: none;
			margin-left: auto; /* Esto lo empuja a la derecha */
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
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>
<div class="main-content"> 
<%if mensajeOk <> "" Then %>
	<div class="message-ok">
		<i class="fas fa-check-circle"></i> <%= mensajeOk %>
	</div>
<%End If %>
<div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;">
    <div style="display: flex; align-items: center; gap: 10px;">
        <h1 style="margin: 0;">Hoja de Ruta: <%= hdrid %></h1>
        <a href="controlhdr.asp?fecha=<%=fecha%>&transportista=<%=transportista%>" 
           class="icon-btn" 
           title="Volver"
           style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
            <i class="fas fa-arrow-left"></i> Volver
        </a>
        <a href="exportar_excel_hdr.asp?hdrid=<%= hdrid %>" 
           class="icon-btn"   
           style="background-color: #28a745; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
           <i class="fa-solid fa-file-csv"></i>Exportar a Excel
        </a>
    </div>
    <a href="terminarhdr.asp?hdrid=<%= hdrid %>&fecha=<%=fecha%>&transportista=<%=transportista%>" 
       class="icon-btn"   
       style="background-color: #ff030e; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
       <i class="fas fa-lock"></i> Cerrar HDR
    </a>
</div>
  
    <h1 style="margin: 0;">Transportista: <%= nombretransportistas %></h1>
    <table class="table table-striped">
        <tbody>
            <% 
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

            Set rs = conn.Execute("EXEC [dbo].[usp_Transportista_HojaDeRuta_Control_Resumen] " & hdrid)
 
            Dim currentForma, nuevaForma
            Dim diferencia, detalleTexto, detalleID
            currentForma = ""
            Do Until rs.EOF
                nuevaForma = rs("CondicionAgrupada")
                FormaPago = rs("FormaPago")
                If nuevaForma <> currentForma Then
                    If currentForma <> "" Then
                        Response.Write "<tr class='subtotal'>"
                        Response.Write "<td colspan='2'>Subtotal</td>"
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
                            <th rowspan="2">Cliente</th>
                            <th rowspan="2">Facturas</th>
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
                FacturaID = "Facturas_" & rs("ClienteID") & Facturas

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
                        <td style="background-color: <%= rs("color") %>"><%= rs("ClienteID") & " - " & rs("ClienteNombre") %></td>
                        <td style="background-color: <%= rs("color") %>"><%= rs("TotalFacturas") %></td>
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
                 <a href="controlhdr_validar_cobranza.asp?hdrid=<%= hdrid %>&clienteid=<%= rs("ClienteID") %>&accion=Editar&formapago=<%= Server.URLEncode(FormaPago) %>" 
                   title="Modificar validaci贸n"
                   style="margin-left: 8px; color: #007bff;">
                    <i class="fas fa-pen"></i>
                </a>
				<span title="<%= Server.HTMLEncode(rs("ValidadoMensaje")) %>" class="validado-span">
                    <i class="fas fa-check-circle" style="color: green;"></i> Revisado
                </span>

        <% 
            ElseIf valEstado = 0 Then
                ' Mostrar como No Validado + bot贸n de editar
        %>
                <a href="controlhdr_validar_cobranza.asp?hdrid=<%= hdrid %>&clienteid=<%= rs("ClienteID") %>&accion=Editar&formapago=<%= Server.URLEncode(FormaPago) %>" 
                   title="Modificar validaci贸n"
                   style="margin-left: 8px; color: #007bff;">
                    <i class="fas fa-pen"></i>
                </a>
                <span title="<%= Server.HTMLEncode(rs("ValidadoMensaje")) %>" class="validado-span" style="color: red;">
                    <i class="fas fa-times-circle"></i> Revisado
                </span>

        <%
            Else
                ' Validado no es ni 1 ni 0: mostrar acciones para validar/invalidar
        %>
                <a href="controlhdr_validar_cobranza.asp?hdrid=<%= hdrid %>&clienteid=<%= rs("ClienteID") %>&accion=Invalidar&formapago=<%= Server.URLEncode(FormaPago) %>"
                   title="Invalidar cobranza" 
                   style="color: red;">
                    <i class="fas fa-times-circle"></i>
                </a>

                <a href="controlhdr_validar_cobranza.asp?hdrid=<%= hdrid %>&clienteid=<%= rs("ClienteID") %>&accion=Validar&formapago=<%= Server.URLEncode(FormaPago) %>" 
                   title="Validar cobranza" 
                   style="color: green; margin-left: 8px;">
                    <i class="fas fa-check-circle"></i>
                </a>
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
                Response.Write "<td colspan='2'>Subtotal</td>"
                Response.Write "<td>$" & FormatNumber(subtotalCobrar, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalCobrado, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalEfectivo, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalCheque, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalTransferencia, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalDescuentos, 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(subtotalDiferencia, 2) & "</td>"
                Response.Write "<td></td><td></td><td></td><td></td></tr></tbody></table><br>"
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
</div>
</body>
</html>
