<%
' Inicia o continua la sesión
Session.Timeout = 20

' Si no existe la variable, la creamos
If Session("Nombre") = "" Then
    Session("Nombre") = "Transportista de Prueba"
End If
%>

<!DOCTYPE html>
<html>
<head><title>Prueba de Sesión</title></head>
<body>
    <h1>Prueba de sesión</h1>
    <p>Nombre en sesión: <strong><%= Session("Nombre") %></strong></p>
    <p><a href="test_sesion.asp">Recargar página</a></p>
</body>
</html>
