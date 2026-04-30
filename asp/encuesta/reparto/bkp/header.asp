<%
' header.asp
If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>
        <%
        If Request("title") <> "" Then
            Response.Write(Request("title"))
        Else
            Response.Write("Panel Transportista")
        End If
        %>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body {
            margin: 0;
            display: flex;
        }

        /* Sidebar */
        #sidebar {
            width: 250px;
            background-color: #34495e;
            color: white;
            height: 100vh;
            position: fixed;
            top: 0;
            left: 0;
            overflow-y: auto;
            padding-top: 60px; /* espacio para header */
        }

        /* Contenido principal */
        #main {
            margin-left: 10px;
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }

        /* Header */
        header {
            background-color: #2c3e50;
            color: white;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: fixed;
            top: 0;
            left: 250px;
            right: 0;
            z-index: 1000;
            height: 60px;
        }

        .menu-toggle {
            display: none;
            background: none;
            border: none;
            color: white;
            font-size: 22px;
            cursor: pointer;
            margin-right: 10px;
        }

        .transportista-nombre {
            font-weight: bold;
            font-size: 16px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            flex: 1;
        }

        .logout {
            background-color: #e74c3c;
            color: white;
            border: none;
            padding: 8px 14px;
            cursor: pointer;
            border-radius: 4px;
            font-size: 14px;
        }

        @media (max-width: 768px) {
            #sidebar {
                left: -250px;
                transition: left 0.3s;
            }

            #sidebar.open {
                left: 0;
            }

            #main {
                margin-left: 0;
            }

            header {
                left: 0;
            }

            .menu-toggle {
                display: inline;
            }
        }

        /* Espacio para que el contenido no quede debajo del header */
        .main-content {
            padding: 80px 20px 20px 20px;
        }
		
.responsive-table {
    width: 100%;
    overflow-x: auto;
    margin: 20px auto;
}
.responsive-table table {
    border-collapse: collapse;
    width: 100%;
    min-width: 600px;
    background: #fff;
}
.responsive-table th, .responsive-table td {
    border: 1px solid #ccc;
    padding: 10px;
    text-align: center;
}
.responsive-table th {
    background-color: #2c3e50;
    color: white;
}
@media (max-width: 768px) {
    .responsive-table table {
        font-size: 14px;
    }
}		

/* Contenedor con scroll horizontal */
.table-wrapper {
  width: 100%;
  overflow-x: auto;
  -webkit-overflow-scrolling: touch; /* para scroll suave en iOS */
  margin-top: 20px;
}

/* Tabla con ancho mínimo */
.table-wrapper table {
  border-collapse: collapse;
  width: 100%;
  min-width: 600px; /* el mínimo ancho que consideres aceptable */
  background: white;
}

.table-wrapper th, .table-wrapper td {
  border: 1px solid #ccc;
  padding: 10px;
  text-align: center;
}

.table-wrapper th {
  background-color: #2c3e50;
  color: white;
}

/* Opcional: Fuente más pequeña en pantallas pequeñas */
@media (max-width: 480px) {
  .table-wrapper table {
    font-size: 13px;
  }
}

    </style>
    <script>
        function toggleSidebar() {
            document.getElementById("sidebar").classList.toggle("open");
        }
    </script>
</head>
<body>

    <!-- Sidebar -->
    <div id="sidebar">
        <ul style="list-style:none; padding:0; margin:0;">
            <li><a href="home.asp" style="color:white; display:block; padding:10px;">🏠 Inicio</a></li>
            <li><a href="rutas.asp" style="color:white; display:block; padding:10px;">📦 Hojas de Ruta</a></li>
            <li><a href="rechazos.asp" style="color:white; display:block; padding:10px;">❌ Rechazos</a></li>
        </ul>
    </div>

    <!-- Contenido principal -->
    <div id="main">
        <header>
            <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
            <div class="transportista-nombre">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></div>
            <form method="post" action="logout.asp" style="margin: 0;">
                <input type="submit" value="Cerrar sesión" class="logout">
            </form>
        </header>

        <div class="main-content">
