<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

'========================================================
' Evitar cache del navegador
'========================================================
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

'========================================================
' Limpiar variables de sesión del portal
'========================================================
Session.Contents.Remove("Empleado_UsuarioID")
Session.Contents.Remove("Empleado_EmpleadoID")
Session.Contents.Remove("Empleado_Usuario")
Session.Contents.Remove("Empleado_RolID")
Session.Contents.Remove("Empleado_Legajo")
Session.Contents.Remove("Empleado_Nombre")
Session.Contents.Remove("Empleado_DebeCambiarClave")

' Variables temporales del login
Session.Contents.Remove("Tmp_UsuarioID")
Session.Contents.Remove("Tmp_EmpleadoID")
Session.Contents.Remove("Tmp_Usuario")
Session.Contents.Remove("Tmp_RolID")
Session.Contents.Remove("Tmp_Legajo")
Session.Contents.Remove("Tmp_Nombre")
Session.Contents.Remove("Tmp_DebeCambiarClave")

'========================================================
' Destruir completamente la sesión
'========================================================
Session.Abandon
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Cerrando sesión</title>
<meta http-equiv="refresh" content="1;url=empleado_login.asp">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<link rel="stylesheet" href="estilos.css">

<style>
body{
    margin:0;
    font-family:Arial, Helvetica, sans-serif;
    background:#f4f6f9;
}
.logout-wrap{
    min-height:100vh;
    display:flex;
    align-items:center;
    justify-content:center;
}
.logout-card{
    background:#fff;
    border:1px solid #ddd;
    border-radius:12px;
    padding:30px;
    text-align:center;
    box-shadow:0 2px 10px rgba(0,0,0,.06);
}
.logout-card h1{
    margin:0 0 10px 0;
    font-size:24px;
}
.logout-card p{
    color:#666;
    margin:0;
}
</style>
</head>

<body>

<div class="logout-wrap">
    <div class="logout-card">
        <h1>Sesión finalizada</h1>
        <p>Redirigiendo al login...</p>
    </div>
</div>

</body>
</html>