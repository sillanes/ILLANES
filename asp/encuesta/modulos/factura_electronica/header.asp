<%@ Language=VBScript %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Session.CodePage = 65001

' Validar sesión
If Session("currentUser") = "" Then
    Response.Redirect "../../login.asp"
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>ILLANES HNOS SRL - Factura Electrónica</title>
    
    <!-- CSS Base -->
    <link rel="stylesheet" type="text/css" href="../../includes/style.css">
    <link rel="stylesheet" type="text/css" href="../../includes/css/new-style.css">
    
    <!-- CSS del Módulo -->
    <link rel="stylesheet" type="text/css" href="css/factura_electronica.css">
</head>
<body>
    <!-- HEADER DEL MÓDULO -->
    <div class="modulo-header">
        <div class="modulo-header-left">
            <h1 class="modulo-titulo">FACTURA ELECTRÓNICA</h1>
        </div>
        <div class="modulo-header-right">
            <span class="usuario-info">Usuario: <strong><%=UCASE(Session("currentUser"))%></strong></span>
            <a href="../../menu.asp" class="btn-volver">← Volver al Menú</a>
        </div>
    </div>

    <!-- CONTENEDOR PRINCIPAL -->
    <div class="modulo-container">
