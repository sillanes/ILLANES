<%
Sub RenderSectionStart(titulo)
%>
    <div class="menu-section">
        <div class="menu-section-title"><%=titulo%></div>
        <div class="menu-grid">
<%
End Sub

Sub RenderSectionEnd()
%>
        </div>
    </div>
<%
End Sub

Sub RenderItem(titulo, url, icono, color)
%>
    <a class="menu-item-card <%=color%>" href="<%=url%>" title="<%=titulo%>">
        <div class="menu-icon">
            <img src="<%=icono%>" alt="<%=titulo%>">
        </div>
        <div class="menu-text"><%=titulo%></div>
    </a>
<%
End Sub
%>
