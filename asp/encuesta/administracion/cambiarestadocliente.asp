<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"
%>
<!--#include file="sidebar.asp" -->

<%
Dim HDR, ClienteID, nuevoEstado, sql, rsCliente
HDR = Trim(Request("hdr"))
ClienteID = Trim(Request("clienteid"))
nuevoEstado = Request.Form("nuevoEstado")

If Request.ServerVariables("REQUEST_METHOD") = "POST" And Len(nuevoEstado) > 0 Then
    sql = "EXEC usp_Transportista_HojaDeRuta_Cliente_Estado_upd " & _
          Request("hdr") & ", " & Request("clienteid") & ", '" & Replace(nuevoEstado, "'", "''") & "'"
    conn.Execute sql
    Response.Write "<div style='color:green;margin-bottom:10px;'>✅ Estado actualizado correctamente.</div>"
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cambiar Estado Cliente</title>
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
            ✅ Estado de cliente actualizado correctamente.
        </div>
    <% End If %>
	
	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		
    <h2>Cambiar Estado Cliente</h2>
   
		
		<a href="/menu.asp" 
		   class="icon-btn" 
		   title="Cerrar hoja de ruta"
		   style="margin-left: left; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
			<i class="fas fa-arrow-left"></i> Menu
		</a>
	</div> 
	
	

    <form method="get" action="cambiarestadocliente.asp" style="margin-bottom:20px;">
        <label>Hoja de Ruta:</label>
        <input type="text" name="hdr" value="<%=HDR%>" />
        &nbsp;&nbsp;
        <label>Cliente ID:</label>
        <input type="text" name="clienteid" value="<%=ClienteID%>" />
        &nbsp;&nbsp;
        <input type="submit" value="Buscar" />
    </form>

<%
If HDR <> "" Or ClienteID <> "" Then
    sql = "EXEC usp_Transportista_HojaDeRuta_Buscar_Analudas " & _
          IIf(HDR = "", "0", HDR) & ", " & IIf(ClienteID = "", "0", ClienteID)

    Set rsCliente = conn.Execute(sql)

    If Not rsCliente.EOF Then
        Dim currentHDR, currentClienteID
        Dim clienteNombre, estadoCliente
        currentHDR = -1
        currentClienteID = -1

        Response.Write "<div class='cards-container'>"

        Do While Not rsCliente.EOF
            If rsCliente("HojaDeRutaID") <> currentHDR Then
                If currentClienteID <> -1 Then
                    Response.Write "</tbody></table>"
                    Call EscribirFormularioMotivo(currentHDR, currentClienteID, estadoCliente)
                    Response.Write "</div>"
                End If
                If currentHDR <> -1 Then Response.Write "</div>"

                currentHDR = rsCliente("HojaDeRutaID")
                Response.Write "<div class='hdr-card'>"
                Response.Write "<h2>Hoja de Ruta ID: " & currentHDR & "</h2>"
                currentClienteID = -1
            End If

            If rsCliente("ClienteID") <> currentClienteID Then
                If currentClienteID <> -1 Then
                    Response.Write "</tbody></table>"
                    Call EscribirFormularioMotivo(currentHDR, currentClienteID, estadoCliente)
                    Response.Write "</div>"
                End If

                currentClienteID = rsCliente("ClienteID")
                clienteNombre = rsCliente("ClienteNombre")
                estadoCliente = rsCliente("Estado")

                Response.Write "<div class='cliente-card'>"
                Response.Write "<h3><i class='fas fa-user'></i> " & clienteNombre & "</h3>"
                Response.Write "<p><strong>Cliente ID:</strong> " & currentClienteID & "</p>"

                Response.Write "<table class='facturas'>"
                Response.Write "<thead><tr>"
                Response.Write "<th>Factura ID</th>"
                Response.Write "<th>Importe Factura</th>"
                Response.Write "<th>Forma Pago</th>"
                Response.Write "<th>Estado</th>"
                Response.Write "</tr></thead><tbody>"
            End If

            Response.Write "<tr>"
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

            Response.Write "</td></tr>"

            rsCliente.MoveNext
        Loop

        Response.Write "</tbody></table>"
        Call EscribirFormularioMotivo(currentHDR, currentClienteID, estadoCliente)
        Response.Write "</div></div>"

    Else
        Response.Write "<p style='color:red;'>❌ No se encontraron clientes.</p>"
    End If

    rsCliente.Close
    Set rsCliente = Nothing
End If

Sub EscribirFormularioMotivo(HDR, ClienteID, FormaActual)
    Response.Write "<form method='post' action='cambiarestadocliente_confirmar.asp' class='update-form'>"
    Response.Write "<input type='hidden' name='hdr' value='" & HDR & "'>"
    Response.Write "<input type='hidden' name='clienteid' value='" & ClienteID & "'>"
    Response.Write "<label for='nuevoEstado_" & ClienteID & "'>Nuevo Estado:</label>"
    Response.Write "<select name='nuevoEstado' id='nuevoEstado_" & ClienteID & "'>"

    Dim rsCondiciones
    Set rsCondiciones = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Estado_Combo")
    Do Until rsCondiciones.EOF
        Dim valorCond, textoCond, sel
        valorCond = rsCondiciones("Valor")
        textoCond = rsCondiciones("Descripcion")
        sel = ""
        If valorCond = FormaActual Then sel = " selected"
        Response.Write "<option value='" & Server.HTMLEncode(valorCond) & "'" & sel & ">" & Server.HTMLEncode(textoCond) & "</option>"
        rsCondiciones.MoveNext
    Loop
    rsCondiciones.Close
    Set rsCondiciones = Nothing

    Response.Write "</select>"
    Response.Write "<button type='submit'><i class='fas fa-save'></i> Actualizar</button>"
    Response.Write "</form>"
End Sub
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
