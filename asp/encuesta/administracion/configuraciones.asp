<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If

' --- Guardar cambios si viene por POST ---
If Request.Form("accion") = "guardar" Then
    Dim i, nombre, valor, total
    total = CInt(Request.Form("totalFilas"))

    For i = 1 To total
        nombre = Request.Form("nombre_" & i)
        valor  = Request.Form("valor_" & i)
        
        If nombre <> "" Then
            sql = "UPDATE Configuracion SET Valor = '" & Replace(valor,"'","''") & "' WHERE Nombre = '" & Replace(nombre,"'","''") & "'"
            Conn.Execute sql
        End If
    Next

    Response.Redirect "configuraciones.asp?msg=ok"
End If

' --- Traer configuraciones ---
sql = "SELECT Nombre, Valor FROM Configuracion ORDER BY Nombre"
Set rs = Conn.Execute(sql)
%>
<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Configuraciones</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    <style>
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 8px; border: 1px solid #ccc; }
        input[type=text] { width: 100%; padding: 5px; }
        .btn { padding: 8px 15px; background:#007bff; color:white; border:none; border-radius:5px; cursor:pointer; }
        .btn:hover { background:#0056b3; }
        .msg { background:#d4edda; color:#155724; padding:10px; margin-bottom:10px; border:1px solid #c3e6cb; }
    </style>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content">
    <% If Request("msg") = "ok" Then %>
        <div class="msg">✅ Configuraciones guardadas correctamente.</div>
    <% End If %>
	
	<h2> Configuraciones <hd/>

    <form method="post">
        <input type="hidden" name="accion" value="guardar">

        <table>
            <tr>
                <th>Nombre</th>
                <th>Valor</th>
            </tr>
            <%
            Dim fila: fila = 0
            Do While Not rs.EOF
                fila = fila + 1
            %>
                <tr>
                    <td>
                        <%= rs("Nombre") %>
                        <input type="hidden" name="nombre_<%=fila%>" value="<%=rs("Nombre")%>">
                    </td>
                    <td>
                        <input type="text" name="valor_<%=fila%>" value="<%=rs("Valor")%>">
                    </td>
                </tr>
            <% 
                rs.MoveNext
            Loop
            %>
        </table>
        <input type="hidden" name="totalFilas" value="<%=fila%>">
        <br>
        <button type="submit" class="btn"><i class="fas fa-save"></i> Guardar Cambios</button>
    </form>
</div>
</body>
</html>
<%
rs.Close: Set rs = Nothing
%>
