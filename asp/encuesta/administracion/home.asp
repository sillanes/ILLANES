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
    <title>Administración - ILLANES HNOS SRL</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        .dashboard-container {
            margin-left: 250px;
            padding: 20px;
            background-color: #f5f5f5;
            min-height: 100vh;
        }

        .dashboard-header {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }

        .dashboard-title {
            font-size: 24px;
            color: #333;
            margin: 0;
        }

        .dashboard-subtitle {
            color: #666;
            margin: 5px 0 0 0;
        }

        .dashboard-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }

        .dashboard-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            transition: transform 0.2s;
        }

        .dashboard-card:hover {
            transform: translateY(-2px);
        }

        .card-icon {
            font-size: 48px;
            color: #2563eb;
            margin-bottom: 15px;
        }

        .card-title {
            font-size: 18px;
            font-weight: bold;
            color: #333;
            margin-bottom: 10px;
        }

        .card-description {
            color: #666;
            line-height: 1.5;
        }

        .card-link {
            display: inline-block;
            margin-top: 15px;
            padding: 8px 16px;
            background: #2563eb;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            transition: background 0.2s;
        }

        .card-link:hover {
            background: #1d4ed8;
            color: white;
            text-decoration: none;
        }

        .stats-container {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
        }

        .stat-item {
            text-align: center;
        }

        .stat-number {
            font-size: 32px;
            font-weight: bold;
            color: #2563eb;
        }

        .stat-label {
            color: #666;
            margin-top: 5px;
        }

        .loading {
            color: #999;
        }

        header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            background: white;
            padding: 18px 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            margin-bottom: 20px;
            position: sticky;
            top: 0;
            z-index: 10;
        }

        header .header-left {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        header .btn-header {
            border: none;
            border-radius: 6px;
            padding: 10px 16px;
            cursor: pointer;
            color: #ffffff;
            background: #2563eb;
            text-decoration: none;
        }

        header .btn-header.red {
            background: #dc2626;
        }
    </style>
</head>
<body>
<header>
    <button class="btn-header" onclick="toggleSidebar()"><i class="fas fa-bars"></i> Menú</button>
    <div class="header-left">
        <strong>👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
        <a href="logout.asp" class="btn-header red"><i class="fas fa-sign-out-alt"></i> Cerrar sesión</a>
    </div>
</header>
    <div class="dashboard-container">
        <div class="dashboard-header">
            <h1 class="dashboard-title">Panel de Administración</h1>
            <p class="dashboard-subtitle">Bienvenido al sistema administrativo de ILLANES HNOS SRL</p>
        </div>

        <div class="stats-container">
            <h3>Resumen de Solicitudes de Alta de Clientes</h3>
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-number loading" id="pendientes-count">-</div>
                    <div class="stat-label">Pendientes</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number loading" id="procesadas-count">-</div>
                    <div class="stat-label">Procesadas Hoy</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number loading" id="total-count">-</div>
                    <div class="stat-label">Total del Mes</div>
                </div>
            </div>
        </div>

        <div class="dashboard-grid">
            <div class="dashboard-card">
                <div class="card-icon">
                    <i class="fas fa-user-plus"></i>
                </div>
                <h3 class="card-title">Alta de Clientes</h3>
                <p class="card-description">
                    Gestionar solicitudes de alta de nuevos clientes enviadas por los vendedores.
                    Revisar documentación y aprobar o rechazar solicitudes.
                </p>
                <a href="altaclientes_admin.asp" class="card-link">
                    <i class="fas fa-arrow-right"></i> Gestionar Altas
                </a>
            </div>

            <div class="dashboard-card">
                <div class="card-icon">
                    <i class="fas fa-cog"></i>
                </div>
                <h3 class="card-title">Configuración</h3>
                <p class="card-description">
                    Configurar parámetros del sistema, usuarios y permisos de acceso.
                </p>
                <a href="configuraciones.asp" class="card-link">
                    <i class="fas fa-arrow-right"></i> Configurar
                </a>
            </div>

            <div class="dashboard-card">
                <div class="card-icon">
                    <i class="fas fa-chart-bar"></i>
                </div>
                <h3 class="card-title">Reportes</h3>
                <p class="card-description">
                    Generar reportes y estadísticas del sistema administrativo.
                </p>
                <a href="#" class="card-link">
                    <i class="fas fa-arrow-right"></i> Ver Reportes
                </a>
            </div>
        </div>
    </div>

    <script>
        // Cargar estadísticas al cargar la página
        document.addEventListener('DOMContentLoaded', function() {
            loadStats();
        });

        function loadStats() {
            // Mostrar loading
            document.getElementById('pendientes-count').textContent = '...';
            document.getElementById('procesadas-count').textContent = '...';
            document.getElementById('total-count').textContent = '...';

            // Hacer petición AJAX para obtener estadísticas
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'get_stats.asp', true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            var stats = JSON.parse(xhr.responseText);
                            document.getElementById('pendientes-count').textContent = stats.pendientes || 0;
                            document.getElementById('procesadas-count').textContent = stats.procesadas_hoy || 0;
                            document.getElementById('total-count').textContent = stats.total_mes || 0;
                        } catch (e) {
                            console.error('Error parsing stats:', e);
                            document.getElementById('pendientes-count').textContent = 'Error';
                            document.getElementById('procesadas-count').textContent = 'Error';
                            document.getElementById('total-count').textContent = 'Error';
                        }
                    } else {
                        console.error('Error loading stats:', xhr.status);
                        document.getElementById('pendientes-count').textContent = 'Error';
                        document.getElementById('procesadas-count').textContent = 'Error';
                        document.getElementById('total-count').textContent = 'Error';
                    }
                }
            };
            xhr.send();
        }
    </script>
</body>
</html>