<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Control</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        .tooltip {
            position: relative;
            display: inline-block;
        }

        .tooltip .tooltiptext {
            visibility: hidden;
            width: 160px;
            background-color: #555;
            color: #fff;
            text-align: center;
            padding: 5px 0;
            border-radius: 6px;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            margin-left: -80px;
            opacity: 0;
            transition: opacity 0.3s;
        }

        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }

        .resumen-totales {
            margin-top: 20px;
            background-color: #f1f1f1;
            padding: 15px;
            border-radius: 8px;
            font-size: 16px;
        }

        .resumen-totales strong {
            color: #333;
        }

        .filtros {
            margin-bottom: 20px;
        }

        .filtros label {
            margin-right: 10px;
        }

        .filtros select,
        .filtros input[type="date"] {
            margin-right: 20px;
        }

        .resaltado {
            background-color: #ffd6e0;
        }
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
<% If Request("msg") = "ok" Then %>
<div style="background:#d4edda;color:#155724;padding:10px;margin-bottom:10px;border:1px solid #c3e6cb;">✅ Hoja de ruta cerrada correctamente.</div>
<% End If %>
	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		
		<h2>Hojas de Ruta Cerradas</h2>
   
		
		<a href="/menu.asp" 
		   class="icon-btn" 
		   title="Cerrar hoja de ruta"
		   style="margin-left: left; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
			<i class="fas fa-arrow-left"></i> Menu
		</a>
	</div> 


<%
Dim fecha, transportista

fecha = Request("fecha")
If fecha = "" Then
    fecha = Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2)
End If

transportista = clng(Request("transportista"))
if transportista="" Then
	transportista=0
End If
%>
    <div class="filtros">
     <form method="get" action="">
        <label for="fecha">Fecha:</label>
        <input type="date" name="fecha" id="fecha" value="<%= fecha %>">

        <label for="transportista">Transportista:</label>
        <select name="transportista" id="transportista">
            <option value="0">Todos</option>
            <% 
            Dim rst
            Set rst = conn.Execute("EXEC usp_Transportista_Sel 0,0 ")
            Do Until rst.EOF
                Dim selected
                If transportista = rst("TransportistaID") Then
                    selected = "selected"
                Else
                    selected = ""
                End If
            %>
            <option value="<%= rst("TransportistaID") %>" <%= selected %>><%= rst("Apellido") & ", " & rst("Nombre")%></option>
            <% 
                rst.MoveNext
            Loop
            rst.Close
            Set rst = Nothing
            %>
        </select>

        <input type="submit" value="Filtrar">
    </form>
    </div>

    <table class="table table-striped table-bordered align-middle">
        <thead class="table-light">
            <tr>
                <th>Hoja de Ruta</th>
                <th>Transportista</th>
                <th>Total a Cobrar</th>
                <th>Total Cobrado</th>
                <th>Acción</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim sqlFiltro
        sqlFiltro = "EXEC [dbo].[usp_Transportista_HojaDeRuta_Cerradas] 0, "

        If fecha <> "" Then
            sqlFiltro = sqlFiltro & " @Fecha = '" & fecha & "'"
        End If

        'If transportista <> "" Then
            If fecha <> "" Then
                sqlFiltro = sqlFiltro & ","
            Else
                sqlFiltro = sqlFiltro & " "
            End If
            sqlFiltro = sqlFiltro & " @TransportistaID = " & transportista
       ' End If
		' response.write sqlFiltro
        Set rs = conn.Execute(sqlFiltro)
        Do Until rs.EOF
            Dim resaltado
            If FormatNumber(rs("TotalACobrar"),0) <> FormatNumber(rs("TotalCobrado"),2) Then
                resaltado = "resaltado"
            Else
                resaltado = ""
            End If
        %>
            <tr class="<%= resaltado %>">
                <td><%= rs("HojaDeRutaID") %></td>
                <td><%= rs("Transportista") %></td>
                <td>$<%= FormatNumber(rs("TotalACobrar"), 2) %></td>
                <td>$<%= FormatNumber(rs("TotalCobrado"), 2) %></td>
                <td><form method="post" action="controlhdr_clientes.asp">
                        <input type="hidden" name="hdrid" value="<%= rs("HojaDeRutaID") %>">
                        <input type="hidden" name="fecha" value="<%=fecha%>">
                        <input type="hidden" name="transportista" value="<%= rs("TransportistaID") %>">
                        <input type="submit" value="Seleccionar">
                    </form>
				</td>
            </tr>
        <%
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
        </tbody>
    </table>
</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>

</body>
</html>
