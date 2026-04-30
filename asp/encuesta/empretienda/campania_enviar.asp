<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then Response.Redirect "../login.asp"

Dim CampaniaID, msg, tipoMsg
CampaniaID = Request.QueryString("CampaniaID")
msg = ""
tipoMsg = ""

If CampaniaID <> "" Then
    On Error Resume Next
    conn.Execute "EXEC usp_Campania_Pendientes_Enviar_Upd " & CampaniaID
    If Err.Number = 0 Then
        msg = "Campaña enviada correctamente."
        tipoMsg = "success"
    Else
        msg = "❌ Error al enviar la campaña: " & Err.Description
        tipoMsg = "error"
    End If
    On Error GoTo 0
Else
    msg = "❌ Campaña no especificada."
    tipoMsg = "error"
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Enviar Campaña</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
body { font-family: Arial, sans-serif; background:#f4f4f4; margin:0; padding:0; }
.main-content { max-width:500px; margin:50px auto; padding:20px; background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.1); text-align:center; }
.msg-success { color:#28a745; background:#d4edda; padding:15px; border-radius:12px; margin-bottom:20px; display:inline-block; font-weight:bold; }
.msg-error { color:#dc3545; background:#f8d7da; padding:15px; border-radius:12px; margin-bottom:20px; display:inline-block; font-weight:bold; }
.btn-volver {
    padding:10px 20px;
    background: linear-gradient(135deg,#007bff,#0056b3);
    color: #fff;
    border:none;
    border-radius:25px;
    cursor:pointer;
    font-weight:bold;
    text-decoration:none;
}
.btn-volver:hover { background: linear-gradient(135deg,#0056b3,#003f88); }
</style>
    <script>
		function toggleSidebar() {
			document.querySelector('.sidebar').classList.toggle('open');
		}

    </script>
</head>
<body>

<!--#include file="header.asp" -->
<div class="main-content">
	<h2>Campaña lista para envío</h2>
    <% If tipoMsg="success" Then %>
        <div class="msg-success"><%=msg%></div>
    <% Else %>
        <div class="msg-error"><%=msg%></div>
    <% End If %>
<br/>
    <a href="campania_activar.asp" class="btn-volver">Volver a Activación</a>
</div>

</body>
</html>
