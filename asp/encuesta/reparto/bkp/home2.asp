<%
' Proteger la página
If Session("nombre") = "" Then
    Response.Redirect "login.asp"
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Inicio Transportista</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- Font Awesome para íconos -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        * {
            box-sizing: border-box;
            font-family: Arial, sans-serif;
        }
        body {
            margin: 0;
            display: flex;
            height: 100vh;
            background-color: #f4f4f4;
        }
        .sidebar {
            width: 220px;
            background-color: #2c3e50;
            color: white;
            padding-top: 20px;
            position: fixed;
            height: 100%;
        }
        .sidebar h2 {
            text-align: center;
            margin-bottom: 30px;
        }
        .sidebar a {
            display: block;
            color: white;
            padding: 12px 20px;
            text-decoration: none;
            transition: background-color 0.3s;
        }
        .sidebar a:hover {
            background-color: #34495e;
        }
        .sidebar i {
            margin-right: 10px;
        }
        .main-content {
            margin-left: 220px;
            padding: 20px;
            width: 100%;
        }
        header {
            background-color: #2c3e50;
            color: white;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        header .left {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        header .logout {
            background-color: #e74c3c;
            color: white;
            border: none;
            padding: 8px 14px;
            cursor: pointer;
            border-radius: 4px;
            font-size: 14px;
        }
        @media (max-width: 768px) {
            .sidebar {
                display: none;
            }
            .main-content {
                margin-left: 0;
            }
        }
    </style>
</head>
<body>

<!-- Menú lateral -->
<div class="sidebar">
    <h2>🚚 Transportista</h2>
    <a href="#"><i class="fas fa-home"></i>Inicio</a>
    <a href="#"><i class="fas fa-map"></i>Hoja de Ruta</a>
    <a href="#"><i class="fas fa-truck"></i>Mis Entregas</a>
    <a href="#"><i class="fas fa-cog"></i>Configuración</a>
</div>

<!-- Contenido principal -->
<div class="main-content">
    <header>
        <div class="left">
            <strong>👤 <%= Session("TransportistaNombre") %></strong>
            <form method="post" action="logout.asp" style="margin: 0;">
                <input type="submit" value="Cerrar sesión" class="logout">
            </form>
        </div>
    </header>

    <h1>Bienvenido al panel del transportista</h1>
    <p>Selecciona una opción del menú para comenzar.</p>
</div>

</body>
</html>
