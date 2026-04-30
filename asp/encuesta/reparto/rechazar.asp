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

Dim latitud, longitud
latitud  = Request.Form("latitud")
longitud = Request.Form("longitud")

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
            .Parameters.Append .CreateParameter("@MotivoID", 3, 1, , motivoID)
            .Parameters.Append .CreateParameter("@texto", 200, 1, 255, textoMotivo)
			.Parameters.Append .CreateParameter("@Latitud",  200, 1, 50, latitud)
			.Parameters.Append .CreateParameter("@Longitud", 200, 1, 50, longitud)
			
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
    <title>Motivo de Rechazo</title>
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

<h2>Motivos</h2>


<% If mensaje <> "" Then %>
<div style="background-color: #f8d7da; color: #721c24; padding: 10px; margin-bottom: 10px; border: 1px solid #f5c6cb;">
    <%= mensaje %>
</div>
<% End If %>


<form method="post" action="rechazar.asp?hdrid=<%= hojaRutaID %>&clienteid=<%= clienteID %>">

<input type="hidden" name="latitud" id="latitud">
<input type="hidden" name="longitud" id="longitud">
<input type="hidden" id="geo_ok" value="0">


    <!-- Grupo: Pendientes -->
    <h3><i class="fas fa-hourglass-half" style="color:#f39c12;"></i> Pendiente</h3>
    <div style="margin-left:15px;">
        <label><input type="radio" name="motivo" value="0" onclick="mostrarCampoTexto()" required> Cerrado</label><br>
        <label><input type="radio" name="motivo" value="1" onclick="mostrarCampoTexto()"> Sin dinero</label><br>
        <label><input type="radio" name="motivo" value="6" onclick="mostrarCampoTexto()"> No se cargó</label>
    </div>

    <br>

    <!-- Grupo: Anulados -->
    <h3><i class="fas fa-times-circle" style="color:#c0392b;"></i> Anular</h3>
    <div style="margin-left:15px;">
        <label><input type="radio" name="motivo" value="2" onclick="mostrarCampoTexto()"> Condición de venta</label><br>
        <label><input type="radio" name="motivo" value="3" onclick="mostrarCampoTexto()"> No lo pidió</label><br>
        <label><input type="radio" name="motivo" value="4" onclick="mostrarCampoTexto()"> Fuera de término</label><br>
        <label><input type="radio" name="motivo" value="5" id="motivo_5" onclick="mostrarCampoTexto()"> Otros</label>
    </div>

    <!-- Campo texto para "Otros" -->
    <div id="campo_texto" style="margin-top: 10px; display: none;">
        <label>Especifique el motivo:<br>
            <textarea name="texto" rows="3" cols="40"></textarea>
        </label>
    </div>

    <br> 
	
	<button type="button"
			id="btnConfirmar"
			style="background-color:#c0392b; color:white;
				   padding:10px 20px; border:none; border-radius:5px;">
		Confirmar
	</button>
	
	
</form>
</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
	
document.getElementById("btnConfirmar").addEventListener("click", function () {

    // Validación básica del motivo
    var motivoSeleccionado = document.querySelector('input[name="motivo"]:checked');
    if (!motivoSeleccionado) {
        alert("Debes seleccionar un motivo.");
        return;
    }

    if (motivoSeleccionado.value === "5") {
        var texto = document.querySelector('textarea[name="texto"]').value.trim();
        if (texto === "") {
            alert("Debes ingresar un motivo cuando seleccionas 'Otros'.");
            return;
        }
    }

    // Si ya tenemos geo → enviar
    if (document.getElementById("geo_ok").value === "1") {
        this.closest("form").submit();
        return;
    }

    if (!navigator.geolocation) {
        alert("Este dispositivo no soporta geolocalización.");
        return;
    }

    navigator.geolocation.getCurrentPosition(
        function (pos) {
            document.getElementById("latitud").value  = pos.coords.latitude.toFixed(6);
            document.getElementById("longitud").value = pos.coords.longitude.toFixed(6);
            document.getElementById("geo_ok").value   = "1";

            // 🔥 ahora sí se envía
            document.getElementById("btnConfirmar").closest("form").submit();
        },
        function () {
            alert(
              "⚠️ No se pudo obtener la ubicación.\n\n" +
              "Si estás en WhatsApp: tocá ⋮ → Abrir en Chrome."
            );
        },
        {
            enableHighAccuracy: true,
            timeout: 15000,
            maximumAge: 0
        }
    );
});	
</script>
</body>
</html>
