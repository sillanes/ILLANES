<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim hdrid, clienteid
hdrid     = Request.QueryString("hdrid")
clienteid = Request.QueryString("clienteid")

If hdrid = "" Or clienteid = "" Then
    Response.Write "Parámetros inválidos."
    Response.End
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Comprobantes del Cliente</title>

<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">

<style>
.table td, .table th { vertical-align: middle; }
.icon-link { font-size: 1.2em; }
.volver-btn {
    background:#2e30c7;
    color:#fff;
    padding:8px 12px;
    border-radius:4px;
    text-decoration:none;
}
.badge {
    background:#28a745;
    color:white;
    padding:4px 8px;
    border-radius:12px;
    font-size:.9em;
}
</style>

<script>
function toggleSidebar() {
    document.querySelector('.sidebar').classList.toggle('open');
}
</script>
</head>

<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()">
        <i class="fas fa-bars"></i>
    </button>
    <strong style="flex:1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin:0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content">

<div style="display:flex; align-items:center; gap:10px; margin-bottom:20px;">
    <h1 style="margin:0;">📎 Comprobantes</h1>
    <a href="javascript:history.back()" class="volver-btn">
        <i class="fas fa-arrow-left"></i> Volver
    </a>
</div>

<table class="table table-striped">
<thead>
<tr>
    <th>#</th>
    <th>Archivo</th>
    <th>Tipo</th>
    <th>Fecha</th>
    <th>Acciones</th>
</tr>
</thead>
<tbody>

<%
Dim sql, rs, contador
contador = 0

Set rs = conn.Execute("EXEC cobranza.Transportista_Comprobante_Sel " & hdrid & "," & clienteid)

If rs.EOF Then
%>
<tr>
    <td colspan="5" style="text-align:center; color:#777;">
        No hay comprobantes cargados para este cliente.
    </td>
</tr>
<%
Else
    Do Until rs.EOF
        contador = contador + 1
%>
<tr>
    <td><%= contador %></td>
    <td>
        <i class="fas fa-file"></i>
        <%= Server.HTMLEncode(rs("NombreArchivo")) %>
    </td>
    <td>
        <span class="badge"><%= UCase(rs("Extension")) %></span>
    </td>
    <td>
        <%= FormatDateTime(rs("FechaSubida"), 2) %>
    </td>
    <td>
        <a class="icon-link" 
           href="<%= "https://illanes-encuesta.ddns.net/comprobantes/" & rs("NombreArchivo")%>" 
           target="_blank"
           title="Ver / Descargar">
            <i class="fas fa-download"></i>
        </a>
    </td>
</tr>
<%
        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing
%>

</tbody>
</table>

</div>
</body>
</html>
