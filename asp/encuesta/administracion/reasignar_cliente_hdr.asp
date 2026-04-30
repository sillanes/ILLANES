<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim hdrID, clienteID, sql, rsCliente, rsHDR,fecha
hdrID = Request.QueryString("HojaDeRutaID")
clienteID = Request.QueryString("ClienteID")
fecha = Request.QueryString("fecha")

' Obtener los datos del cliente y las facturas
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cabecera_selV2 " & hdrID & ", " & clienteID
'response.write sql
Set rsCliente = Conn.Execute(sql)

' Obtener las hojas de ruta abiertas del transportista
Dim  rsHDReasignaciones 
sql = "EXEC dbo.[usp_Transportista_HojaDeRuta_Reasignacion_Resumen] " & hdrID & ", " & clienteID
'response.write sql
Set rsHDReasignaciones = Conn.Execute(sql)

' Obtener las hojas de ruta abiertas del transportista
Dim  rsHDRDisponibles 
sql = "EXEC dbo.[usp_Transportista_HojaDeRuta_Para_Reasignar] 0,0, " &clienteID   
'response.write sql
Set rsHDRDisponibles = Conn.Execute(sql)
%>
<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0"> 
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <title>Reasignar Cliente a otra Hoja de Ruta</title>
    <link rel="stylesheet" href="estilos.css">
    <script>
        function confirmarReasignacion(nuevaHdrID) {
            if (confirm("¿Está seguro que desea reasignar este cliente a la nueva hoja de ruta?")) {
                window.location.href = "procesar_reasignacion.asp?HDROrigen=<%=hdrID%>&ClienteID=<%=clienteID%>&HDRDestino=" + nuevaHdrID;
            }
        }
    </script>
    <style>
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 10px; border: 1px solid #ccc; text-align: left; }
        .action-button { background-color: #007bff; color: white; border: none; padding: 5px 10px; cursor: pointer; border-radius: 5px; }
    </style>
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
<% If Request("msg") = "reasignado" Then %>
    <div style="background-color: #d4edda; color: #155724; padding: 10px; margin-bottom: 10px; border: 1px solid #c3e6cb;">
        ✅ Cliente reasignado correctamente.
    </div>
<% End If %>


	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		
	<h1 style="margin: 0;">Reasignar Pendientes </h1>
		<a href="reasignar.asp?fecha=<%=fecha%>"
		   class="icon-btn" 
		   title="Volver"
		   style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none; margin-left: 10px;">
			<i class="fas fa-arrow-left"></i> Volver
		</a>
	</div> 
	
 
    <h3>Datos del Cliente</h3>
    <table>
        <tr><th>ClienteID</th><th>Cliente Nombre</th><th>Cantidad Facturas</th><th>Importe</th></tr>
        <% Do While Not rsCliente.EOF %>
            <tr>
                <td><%= rsCliente("ClienteID") %></td>
                <td><%= rsCliente("ClienteNombre") %></td>
                <td><%= rsCliente("TotalFacturas") %></td>
                <td>$ <%= FormatNumber(rsCliente("ImporteACobrar"), 2) %></td>
            </tr>
        <% rsCliente.MoveNext: Loop %>
    </table>
    
    <h3>Historial de Reasignaciones</h3>
    <table>
        <tr>
            <th>Hoja de Ruta Origen</th>
            <th>Hoja de Ruta Destino</th>
            <th>Fecha HDR</th>
            <th>Fecha Reasignación</th>
            <th>Forma de Pago</th>
            <th>Importe Factura</th>
            <th>Estado</th>
        </tr>
        <% 
        If Not rsHDReasignaciones.EOF Then
            Do While Not rsHDReasignaciones.EOF 
                ' Estado con iconos
                Dim estadoIcono
                Select Case rsHDReasignaciones("Estado")
                    Case 0 ' En proceso
                        estadoIcono = "<i class='fas fa-sync fa-spin' style='color:blue;' title='En Proceso'></i> En Proceso"
                    Case 2 ' Pendiente
                        estadoIcono = "<i class='fas fa-clock' style='color:orange;' title='Pendiente'></i> Pendiente"
                    Case 3 ' Pendiente
                        estadoIcono = "<i class='fas fa-exclamation-circle' style='color:red;' title='Pendiente'></i> Pendiente"
                    Case Else
                        estadoIcono = rsHDReasignaciones("Estado")
                End Select
        %>
            <tr>
                <td><%= rsHDReasignaciones("HojaDeRutaOrigen") %></td>
                <td><%= rsHDReasignaciones("HojaDeRutaDestino") %></td>
                <td><%= FormatearFecha(rsHDReasignaciones("FechaHDR")) %></td>
                <td><%= FormatearFechaHora(rsHDReasignaciones("FechaReasignacion")) %></td>
                <td><%= rsHDReasignaciones("FormaPago") %></td>
                <td>$ <%= FormatNumber(rsHDReasignaciones("ImporteFactura"), 2) %></td>
                <td><%= estadoIcono %></td>
            </tr>
        <%  
            rsHDReasignaciones.MoveNext
            Loop 
        Else
        %>
            <tr><td colspan="7" style="text-align:center;">No hay reasignaciones registradas.</td></tr>
        <% End If %>
    </table>



    <h3>Seleccionar nueva Hoja de Ruta</h3>
    <table>
        <tr><th>HojaDeRutaID</th><th>Fecha</th><th>Zona</th><th>Transportista</th><th>Acción</th></tr>
        <% Do While Not rsHDRDisponibles.EOF %>
            <tr>
                <td><%= rsHDRDisponibles("HojaDeRutaID") %></td>
                <td><%= FormatDateTime(rsHDRDisponibles("DateAdded"),2) %></td>
                <td><%= rsHDRDisponibles("Zona") %></td>
                <td><%= rsHDRDisponibles("transportista") %></td>
                <td>
					<% If NOT rsHDRDisponibles("Reasignado") = 1 Then %>
                    <button class="action-button" onclick="confirmarReasignacion(<%= rsHDRDisponibles("HojaDeRutaID") %>)">
                        Reasignar
                    </button>
					<%else%>
					<a href="#" class="icon-btn" title="Reasignado"><i class="fas fa-ban"></i></a>
					<%End if%>
					
                </td>
            </tr>
        <% rsHDRDisponibles.MoveNext: Loop %>
    </table>
</div>
</body>
</html>
<%
rsCliente.Close: Set rsCliente = Nothing
rsHDRDisponibles.Close: Set rsHDRDisponibles = Nothing
rsHDReasignaciones.Close: Set rsHDRDisponibles = Nothing
 
Function FormatearFecha(fecha)
    If IsNull(fecha) Or fecha = "" Then
        FormatearFecha = ""
    Else
        FormatearFecha = Right("0" & Day(fecha),2) & "/" & _
                         Right("0" & Month(fecha),2) & "/" & _
                         Year(fecha)
    End If
End Function

Function FormatearFechaHora(fecha)
    If IsNull(fecha) Or fecha = "" Then
        FormatearFechaHora = ""
    Else
        FormatearFechaHora = Right("0" & Day(fecha),2) & "/" & _
                             Right("0" & Month(fecha),2) & "/" & _
                             Year(fecha) & " " & _
                             Right("0" & Hour(fecha),2) & ":" & _
                             Right("0" & Minute(fecha),2) & ":" & _
                             Right("0" & Second(fecha),2)
    End If
End Function
%>


%>
