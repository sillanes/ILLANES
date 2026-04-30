<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "../login.asp"

' Obtener templates de WhatsApp
Dim rsTemplates, sql
Set rsTemplates = Server.CreateObject("ADODB.Recordset")
sql = "EXEC usp_WhastAPP_Templates"
rsTemplates.Open sql, Conn

Dim templates : templates = ""
If Not rsTemplates.EOF Then
    templates = "<option value=''>-- Seleccione un template --</option>"
    Do Until rsTemplates.EOF
        templates = templates & "<option value='" & rsTemplates("Nombre") & "'>" &  rsTemplates("Texto") & "</option>"
        rsTemplates.MoveNext
    Loop
End If
rsTemplates.Close
Set rsTemplates = Nothing
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Creación de Campañas</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="ventas.css">
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

		function toggleSidebar() {
			document.querySelector('.sidebar').classList.toggle('open');
		}

    </script>
</head>
<body>


<!--#include file="header.asp" -->
<div class="main-content">
<h2>Configurar Campaña</h2>
<form method="post" action="guardar_campania_datos.asp" class="form-campania">
    <label>Nombre de la campaña</label>
    <input type="text" name="NombreCampania" required>
    <label>Descripción</label>
    <textarea name="Descripcion"></textarea>
    <label>Canal</label>
    <input type="radio" name="Canal" value="WhatsApp" onclick="toggleChannels()"> WhatsApp
    <input type="radio" name="Canal" value="Email" onclick="toggleChannels()"> Email

    <div id="whatsappBox" style="display:none;">
        <label>Mensaje WhatsApp</label>
        <textarea name="MensajeWhatsApp"></textarea>
        <label>Template WhatsApp</label>
        <select name="TemplateWhatsApp"><%=templates%></select>
    </div>



	<br/>
    <button type="submit" class="btn">Guardar Campaña</button>
</form>
</div>
</body>
</html>
