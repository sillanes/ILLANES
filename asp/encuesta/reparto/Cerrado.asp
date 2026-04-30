<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

Dim hojaRutaID, clienteID, motivoID, textoMotivo
hojaRutaID = Request.QueryString("hdrid")
clienteID = Request.QueryString("clienteid")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    motivoID = Request.Form("motivo")
    textoMotivo = Request.Form("texto")

    If motivoID = "5" And textoMotivo = "" Then
        mensaje = "Debes ingresar un motivo cuando seleccionas 'Otros'."
    Else
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        With cmd
            .ActiveConnection = conn
            .CommandType = 4 ' Stored Procedure
            .CommandText = "[dbo].[usp_Transportista_HojaDeRuta_Cliente_Cerrar]"
            .Parameters.Append .CreateParameter("@HojaDeRutaID", 3, 1, , hojaRutaID)
            .Parameters.Append .CreateParameter("@ClienteID", 3, 1, , clienteID)
            .Parameters.Append .CreateParameter("@MotivioID", 3, 1, , motivoID)
            .Parameters.Append .CreateParameter("@texto", 200, 1, 255, textoMotivo)
            .Execute
        End With
        Set cmd = Nothing

        Response.Redirect "hojaderutaV3.asp?hdrid=" & hojaRutaID
    End If
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Comercio cerrado</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <script>
    function mostrarCampoTexto() {
        var otros = document.getElementById('motivo_5');
        var campoTexto = document.getElementById('campo_texto');
        campoTexto.style.display = otros.checked ? 'block' : 'none';
    }
    </script>
    <style>
        .tooltip {
            position: relative;
            display: inline-block;
        }

        .tooltip .tooltiptext {
            visibility: hidden;
            width: 200px;
            background-color: #333;
            color: #fff;
            text-align: center;
            border-radius: 5px;
            padding: 5px;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            margin-left: -100px;
            opacity: 0;
            transition: opacity 0.3s;
        }

        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
    </style>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content"> 

  <div class="d-flex justify-content-between align-items-center mb-4"> 
    <a href="hojaderutaV3.asp?hdrid=<%= hojaRutaID %>"  
       class="icon-btn" 
       title="Volver"
       style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none; margin-right: 10px;">
        <i class="fas fa-arrow-left"></i> Volver
    </a> 
  </div>

<h2>Selecciona el motivo de rechazo</h2>


<% If mensaje <> "" Then %>
<div style="background-color: #f8d7da; color: #721c24; padding: 10px; margin-bottom: 10px; border: 1px solid #f5c6cb;">
    <%= mensaje %>
</div>
<% End If %>


<form method="post" action="rechazar.asp?hdrid=<%= hojaRutaID %>&clienteid=<%= clienteID %>">
    <div>
        <label><input type="radio" name="motivo" value="0" onclick="mostrarCampoTexto()" required> Cerrado</label><br>
        <label><input type="radio" name="motivo" value="1" onclick="mostrarCampoTexto()"> Sin dinero</label><br>
        <label><input type="radio" name="motivo" value="2" onclick="mostrarCampoTexto()"> Condición de venta</label><br>
        <label><input type="radio" name="motivo" value="3" onclick="mostrarCampoTexto()"> No lo pidió</label><br>
        <label><input type="radio" name="motivo" value="4" onclick="mostrarCampoTexto()"> Fuera de término</label><br> 
        <label><input type="radio" name="motivo" value="5" id="motivo_5" onclick="mostrarCampoTexto()"> Otros</label>
    </div>
    <div id="campo_texto" style="margin-top: 10px; display: none;">
        <label>Especifique el motivo:<br>
            <textarea name="texto" rows="3" cols="40"></textarea>
        </label>
    </div>
    <br>
    <input type="submit" value="Confirmar rechazo" style="background-color: #c0392b; color: white; padding: 10px 20px; border: none; border-radius: 5px;">
	
</form>
</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>
</body>
</html>
