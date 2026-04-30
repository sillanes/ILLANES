<%@ Language="VBScript" %>
 
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim mensaje
mensaje = Request.QueryString("mensaje")
If mensaje = "" Then
    mensaje = "Ha ocurrido un error inesperado al registrar el pago."
Else
    mensaje = Server.HTMLEncode(mensaje)
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Entrega de Facturas</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="estilos.css"> 
</head>
<body class="bg-light">

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>
<div class="container mt-5">
  <div class="alert alert-danger p-4 rounded shadow">
    <h4 class="alert-heading">Error al registrar el pago</h4>
    <p><%= mensaje %></p>
    <hr>
    <a href="javascript:history.back()" class="btn btn-secondary">Volver</a>
  </div>
</div>


</body>
</html> 
