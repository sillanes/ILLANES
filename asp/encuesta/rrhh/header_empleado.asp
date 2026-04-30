<header>
    <button class="menu-toggle" onclick="toggleSidebar()">
        <i class="fas fa-bars"></i>
    </button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>

    <div class="header-buttons">
        <form method="post" action="../rrhh/empleado_logout.asp" style="margin:0;">
            <button type="submit" class="btn-header red">
                <i class="fas fa-sign-out-alt"></i> Cerrar sesión
            </button>
        </form>		
    </div>
</header>