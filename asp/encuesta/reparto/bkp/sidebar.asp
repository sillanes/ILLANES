<div class="sidebar" id="sidebar">
    <h2>🚚 Transportista</h2>
    <a href="home.asp"><i class="fas fa-home"></i>Inicio</a>
    <a href="hojaderuta.asp"><i class="fas fa-map"></i>Hoja de Ruta</a> 
</div>
<style>
    .sidebar {
        width: 220px; background-color: #2c3e50; color: white;
        position: fixed; height: 100%; top: 0; left: 0;
        padding-top: 20px; transition: left 0.3s ease;
        z-index: 1000;
    }
    .sidebar h2 { text-align: center; margin-bottom: 30px; }
    .sidebar a {
        display: block; color: white; padding: 12px 20px;
        text-decoration: none;
    }
    .sidebar a:hover { background-color: #34495e; }
    .sidebar i { margin-right: 10px; }

    @media (max-width: 768px) {
        .sidebar {
            left: -220px;
        }
        .sidebar.open {
            left: 0;
        }
    }
</style>
