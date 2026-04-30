<%
' Sidebar reutilizable para el módulo Factura Electrónica
Dim currentPage
currentPage = Request.ServerVariables("SCRIPT_NAME")
%>

<div class="modulo-sidebar">
    <nav class="sidebar-nav">
        <ul class="sidebar-menu">
            <li class="sidebar-item <%if InStr(currentPage, "enviarwhatsapp") > 0 then response.write("active") end if%>">
                <a href="enviarwhatsapp.asp" class="sidebar-link">
                    <img src="../../images/whastapplogo2.png" alt="icon" class="sidebar-icon">
                    <span>Envío WhatsApp</span>
                </a>
            </li>
            
            <li class="sidebar-item <%if InStr(currentPage, "reenviowhatsapp") > 0 then response.write("active") end if%>">
                <a href="reenviowhatsapp.asp" class="sidebar-link">
                    <img src="../../images/reenviar.png" alt="icon" class="sidebar-icon">
                    <span>Reenvío WhatsApp</span>
                </a>
            </li>
            
            <li class="sidebar-item <%if InStr(currentPage, "facturadigital") > 0 then response.write("active") end if%>">
                <a href="facturadigital.asp" class="sidebar-link">
                    <img src="../../images/alta.png" alt="icon" class="sidebar-icon">
                    <span>Adherir Cliente</span>
                </a>
            </li>
            
            <li class="sidebar-item <%if InStr(currentPage, "facturadigitalpendientes") > 0 then response.write("active") end if%>">
                <a href="facturadigitalpendientes.asp" class="sidebar-link">
                    <img src="../../images/reporte.png" alt="icon" class="sidebar-icon">
                    <span>Mensajes con Error</span>
                </a>
            </li>
        </ul>
    </nav>
</div>
