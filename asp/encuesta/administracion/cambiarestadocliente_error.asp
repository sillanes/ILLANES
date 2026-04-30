<%
Dim mensaje
mensaje = Request.QueryString("msg")
%>

<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"
%>
<!--#include file="sidebar.asp" -->

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Error en la operación</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />	
    <style>
        body { font-family: Arial, sans-serif; background: #fff4f4; padding: 0px; }
        .error-box {
            border: 1px solid #f5c6cb;
            background-color: #f8d7da;
            color: #721c24;
            padding: 20px;
            border-radius: 5px;
            max-width: 600px;
            margin: auto;
        }
        a { color: #004085; text-decoration: none; }
        a:hover { text-decoration: underline; }
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
    
    <div class="error-box">
        <h2>❌ Error</h2>
        <p><%= Server.HTMLEncode(mensaje) %></p>
        <p><a href="javascript:history.back()">⬅ Volver</a></p>
    </div>

</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>


</body>
</html>
