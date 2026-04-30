<div class="sidebar"> 
    <ul class="menu"> 
        <li class="menu-item">
            <span><i class="fa-solid fa-gear"></i> Recibos</span>
            <ul class="submenu">
                <li><a href="rrhh_recibos.asp"><i class="fa-solid fa-magnifying-glass"></i> Recibos</a></li> 
                <li><a href="rrhh_recibos_carga.asp"><i class="fa-solid fa-plus"></i> Cargar</a></li> 
                <li><a href="rrhh_recibos_lotes.asp"><i class="fa-solid fa-file-lines"></i> Lotes</a></li> 
                <li><a href="rrhh_busqueda.asp"><i class="fa-solid fa-file-arrow-down"></i> Busqueda</a></li> 
            </ul>
        </li>
        <li class="menu-item">
            <span><i class="fa-solid fa-bell"></i> Notificaciones</span>
            <ul class="submenu">
                <li><a href="rrhh_notificaciones_masivas.asp"><i class="fa-solid fa-envelopes-bulk"></i> Notificar Lote</a></li>  
                <li><a href="rrhh_notificaciones.asp"><i class="fa-solid fa-person-circle-exclamation"></i> Notificar</a></li>  
            </ul>
        </li> 

        <li class="menu-item">
            <span><i class="fa-solid fa-person"></i> Personal</span>
            <ul class="submenu">
                <li><a href="rrhh_empleados.asp"><i class="fa-solid fa-users"></i> Nomina</a></li>  
                <li><a href="rrhh_vacaciones.asp"><i class="fa-solid fa-suitcase"></i> Vacaciones</a></li>  
                <li><a href="rrhh_vacaciones_dashboard.asp"><i class="fa-solid fa-plane-lock"></i> Vac. Dashboard</a></li>  
            </ul>
        </li>
        <li class="menu-item">
            <span><i class="fa-brands fa-internet-explorer"></i> Web</span>
            <ul class="submenu">
                <li><a href="rrhh_blanqueo_usuario.asp"><i class="fa-regular fa-user"></i> Blanqueo Usuario</a></li>  
            </ul>
        </li>	
		
    </ul>
</div>

<style>
:root {
    --sidebar-width: 250px;
}

/* Sidebar fijo */
.sidebar {
    width: var(--sidebar-width);
    background: #2c3e50;
    color: #ecf0f1;
    height: 100vh;
    position: fixed;
    top: 0;
    left: 0;
    padding-top: 20px;
    overflow-y: auto;
    z-index: 1000;
    box-sizing: border-box;
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
    box-sizing: border-box;
}

/* Texto item principal */
.sidebar ul.menu > li > span {
    display: block;
    font-weight: bold;
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

/* Contenido principal */
.main-content {
    margin-left: var(--sidebar-width);
    width: calc(100% - var(--sidebar-width));
    padding: 20px;
    box-sizing: border-box;
    min-height: 100vh;
}

/* Responsive */
@media (max-width: 768px) {
    .sidebar {
        width: 100%;
        height: auto;
        position: relative;
        padding: 10px;
    }

    .sidebar ul.menu > li {
        padding: 8px 10px;
    }

    .main-content {
        margin-left: 0;
        width: 100%;
        padding: 15px;
    }
}
</style>