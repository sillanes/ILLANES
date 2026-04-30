<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"

Dim HDR, ClienteID, nuevoEstado, estadoClienteActual, rs
HDR = Trim(Request.Form("hdr"))
ClienteID = Trim(Request.Form("clienteid"))
nuevoEstado = Trim(Request.Form("nuevoEstado")) 
Select Case Trim(Request.Form("nuevoEstado"))
	Case 1
		nuevoEstadoLeyenda = "Entregado"
	Case 2
		nuevoEstadoLeyenda = "Anulado"
	Case 3
		nuevoEstadoLeyenda = "Pendiente"
	Case Else
		nuevoEstadoLeyenda = "(No encontrado)"
End Select
			

' Obtener forma de pago actual desde la base
Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Buscar_Analudas " & HDR & ", " & ClienteID)

If Not rs.EOF Then
estadoClienteActual = rs("Estado")
            Select Case rs("Estado")
                Case 1
                    estadoClienteActualLeyenda = "Entregado"
                Case 2
                    estadoClienteActualLeyenda = "Anulado"
                Case 3
                    estadoClienteActualLeyenda = "Pendiente"
                Case Else
                    estadoClienteActualLeyenda = "(No encontrado)"
            End Select
			
Else
    estadoClienteActualLeyenda = "(No encontrado)"
End If
rs.Close
Set rs = Nothing
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <title>Confirmar Cambio</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" /> 
    <style>
        .box {
            max-width: 500px;
            margin: auto;
            padding: 20px;
            background: #f7f7f7;
            border: 1px solid #ccc;
            border-radius: 8px;
        }
        .box h2 {
            margin-top: 0;
        }
        .box label {
            display: block;
            margin-top: 10px;
        }
        .box textarea {
            width: 100%;
            height: 80px;
        }
        .box button {
            margin-top: 15px;
            background-color: #3498db;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
        }
		
		
		
		.boton-container {
			display: flex;
			justify-content: space-between;
			align-items: center;
			margin-top: 30px;
			gap: 20px;
		}

		.boton-container form,
		.boton-container .boton-volver {
			flex: 1;
		}

		.boton-container button,
		.boton-container a {
			display: block;
			width: 100%;
			text-align: center;
			background-color: #3498db;
			color: white;
			padding: 12px 0;
			border: none;
			border-radius: 6px;
			font-size: 16px;
			font-weight: bold;
			text-decoration: none;
			transition: background-color 0.3s ease;
		}

		.boton-container button:hover,
		.boton-container a:hover {
			background-color: #2c80b4;
		}

		.boton-container i {
			margin-right: 6px;
		}

.form-cambio-estado .acciones {
    display: flex;
    gap: 20px;
    margin-top: 16px;
}

.form-cambio-estado .acciones .btn {
    display: flex;               /* Para centrar ícono y texto */
    align-items: center;         /* Centrado vertical */
    justify-content: center;     /* Centrado horizontal */
    width: 150px;                 /* Mismo ancho */
    height: 48px;                 /* Misma altura */
    padding: 0;
    border-radius: 6px;
    font-size: 16px;
    font-weight: bold;
    text-decoration: none;
    border: none;
    background: #3498db;
    color: #fff;
    transition: background-color 0.3s ease;
    cursor: pointer;
}

.form-cambio-estado .acciones .btn i {
    font-size: 16px;   /* Igual tamaño de ícono en ambos */
    margin-right: 6px;
    display: inline-block;
}

.form-cambio-estado .acciones .btn:hover {
    background-color: #2c80b4;
}

.form-cambio-estado .acciones .btn {
  margin-top: 0 !important;   /* anula el margin-top heredado del .box button */
}
@media (max-width: 520px) {
    .form-cambio-estado .acciones {
        flex-direction: column;
    }
}

 
		
		
    </style>
</head>
<body>


<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout" />
    </form>
</header>


<div class="main-content">
<div class="box">
    <h2>Confirmar Cambio Estado</h2> 
	
        <p><strong>Hoja de Ruta:</strong> <%=HDR%></p>
        <p><strong>Cliente ID:</strong> <%=ClienteID%></p>
        <p><strong>Estado Actual:</strong> <%=estadoClienteActualLeyenda%></p>
        <p><strong>Nuevo Estado:</strong> <%=nuevoEstadoLeyenda%></p>


        <input type="hidden" name="hdr" value="<%=HDR%>">
        <input type="hidden" name="clienteid" value="<%=ClienteID%>">
        <input type="hidden" name="nuevoEstado" value="<%=nuevoEstado%>">


<form method="post" action="cambiarestadocliente_guardar.asp" class="form-cambio-estado">
    <input type="hidden" name="hdr" value="<%=HDR%>">
    <input type="hidden" name="clienteid" value="<%=ClienteID%>">
    <input type="hidden" name="nuevoEstado" value="<%=nuevoEstado%>">

    <label for="observaciones">Observaciones</label>
    <textarea name="observaciones" id="observaciones" placeholder="Ingrese motivo o comentario..."></textarea>

    <div class="acciones">
        <button type="submit" class="btn confirmar">
            <i class="fas fa-check-circle"></i> Confirmar
        </button>
        <a href="cambiarestadocliente.asp?hdr=<%=Server.URLEncode(HDR)%>" class="btn volver">
            <i class="fas fa-arrow-left"></i> Volver
        </a>
    </div>
</form>


 	 
</div>

</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>


</body>


</html>
