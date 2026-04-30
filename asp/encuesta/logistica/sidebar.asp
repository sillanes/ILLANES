<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<div class="sidebar"> 
    <ul class="menu">
        <li><a href="validarpendientes.asp"><i class="fas fa-route"></i>Pendietes</a></li>  
        <li class="menu-item">
            <span><i class="fa-solid fa-gear"> </i>  Navidad</span>
            <ul class="submenu">
                <li><a href="Ordenes.asp"><i class="fa-solid fa-pen-to-square"></i> Ordenes</a></li> 
						
				<li><a href="Ordenespendientes.asp"><i class="fa-solid fa-truck-fast"></i>Ordenes Pendientes</a></li>
            </ul>
        </li>

        <li class="menu-item">
            <span><i class="fa-solid fa-gear"> </i>  Vehiculos</span>
            <ul class="submenu">
				<li><a href="vehiculos_abm.asp"><i class="fa-solid fa-car"></i> ABM Vehículos</a></li>
				<li><a href="mantenimiento.asp"><i class="fa-solid fa-brush"></i> Mantenimientos</a></li>
				<li><a href="vehiculos_home.asp"><i class="fa-solid fa-car"></i> Verificar Vehiculo</a></li> 	
				<li><a href="vehiculos_dashboard_mensual.asp"><i class="fa-solid fa-truck-fast"></i>Panel Mensual</a></li>
				<li><a href="vehiculos_dashboard_diario.asp"><i class="fa-solid fa-truck-fast"></i>Panel Diario</a></li>
				
            </ul>
        </li>
		
		
		<li><a href="logout.asp"><i class="fas fa-sign-out-alt"></i> Salir</a></li>
		
    </ul>
</div>

<style>
/* Sidebar fijo */
.sidebar {
    width: 220px;
    background: #2c3e50;
    color: #ecf0f1;
    height: 100vh;
    position: fixed;
    padding-top: 20px;
}

/* Título */
.sidebar h2 {
    text-align: center;
    margin-bottom: 20px;
}

/* Lista principal */
.sidebar ul.menu {
    list-style: none;
    padding: 0;
    margin: 0;
}

/* Item principal */
.sidebar ul.menu > li {
    padding: 10px 20px;
}

/* Submenú siempre visible */
.sidebar ul.menu li .submenu {
    display: block;
    list-style: none;
    padding-left: 15px;
    margin-top: 5px;
}

/* Submenú items */
.sidebar ul.menu li .submenu li {
    padding: 8px 0;
}

/* Links */
.sidebar ul.menu li .submenu li a {
    color: #ecf0f1;
    text-decoration: none;
    font-size: 14px;
}

/* Hover */
.sidebar ul.menu li .submenu li a:hover {
    text-decoration: underline;
}

/* Responsive: sidebar colapsa a top */
@media (max-width: 768px) {
    .sidebar {
        width: 100%;
        height: auto;
        position: relative;
        padding: 10px;
    }
    .sidebar ul.menu li {
        padding: 5px 0;
    }
}
</style>
