<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If
 
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
    
</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
 	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		 
   
    <h2>Bienvenido <%= Server.HTMLEncode(Session("currentUser"))%> </h2>
		 
	</div> 
</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>
</body>
</html>


