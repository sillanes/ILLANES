<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "../login.asp"


Dim NombreCampania, Descripcion, Canal
Dim MensajeWhatsApp, TemplateWhatsApp
Dim MensajeEmailAsunto, MensajeEmailCuerpo
Dim cmd

NombreCampania = Request.Form("NombreCampania")
Descripcion = Request.Form("Descripcion")
Canal = Request.Form("Canal")
MensajeWhatsApp = Request.Form("MensajeWhatsApp")
TemplateWhatsApp = Request.Form("TemplateWhatsApp")
MensajeEmailAsunto = Request.Form("MensajeEmailAsunto")
MensajeEmailCuerpo = Request.Form("MensajeEmailCuerpo")

' Guardar en base usando stored procedure
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = Conn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "usp_Campania_Guardar"

cmd.Parameters.Append cmd.CreateParameter("@NombreCampania", 200, 1, 200, NombreCampania)
cmd.Parameters.Append cmd.CreateParameter("@Descripcion", 200, 1, 500, Descripcion)
cmd.Parameters.Append cmd.CreateParameter("@Canal", 200, 1, 50, Canal)
cmd.Parameters.Append cmd.CreateParameter("@MensajeWhatsApp", 200, 1, 4000, MensajeWhatsApp)
cmd.Parameters.Append cmd.CreateParameter("@TemplateWhatsApp", 200, 1, 200, TemplateWhatsApp)
cmd.Parameters.Append cmd.CreateParameter("@MensajeEmailAsunto", 200, 1, 200, MensajeEmailAsunto)
cmd.Parameters.Append cmd.CreateParameter("@MensajeEmailCuerpo", 200, 1, 4000, MensajeEmailCuerpo)
cmd.Parameters.Append cmd.CreateParameter("@UsuarioCreador", 200, 1, 100, Session("currentUser"))

cmd.Execute
Set cmd = Nothing
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Creación de Campañas</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        form label {
            display: block;
            margin-top: 15px;
            font-weight: bold;
            color: #444;
        } 
        form input[type="text"],
        form input[type="file"],
        form textarea,
        form select {
            width: 100%;
            padding: 8px 10px;
            margin-top: 6px;
            border: 1px solid #ccc;
            border-radius: 6px;
            font-size: 14px;
        }

        form textarea {
            min-height: 80px;
            resize: vertical;
        }

        .radio-inline {
            margin-right: 15px;
            margin-top: 10px;
        }

        .btn {
            margin-top: 20px; 
            background-color: #3498db;
            border: none;
            border-radius: 6px;
            color: #fff;
            cursor: pointer;
            font-size: 15px;
            padding: 10px 20px;
        }

        .btn:hover {
            background-color: #2980b9;
        }

        @media (max-width: 768px) {
            .main-content {
                margin-left: 0;
                padding: 15px;
            }

            form {
                width: 100%;
            }
        }
    </style>
    <script>
        function toggleChannels() {
            var whatsappBox = document.getElementById("whatsappBox");
            var emailBox = document.getElementById("emailBox");

            var canal = document.querySelector('input[name="Canal"]:checked');

            if (canal && canal.value === "WhatsApp") {
                whatsappBox.style.display = "block";
                emailBox.style.display = "none";
            } else if (canal && canal.value === "Email") {
                whatsappBox.style.display = "none";
                emailBox.style.display = "block";
            } else {
                whatsappBox.style.display = "none";
                emailBox.style.display = "none";
            }
        }

        window.onload = toggleChannels;
    </script>
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


<div style="background:#d4edda;color:#155724;padding:10px;margin-bottom:10px;border:1px solid #c3e6cb;">✅ Campaña guardada correctamente.</div>


</div>

</body>
</html>