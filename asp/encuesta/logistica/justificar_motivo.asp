<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"

Dim HDR, ClienteID, FacturaID, mensajeOk, mensajeError
HDR = Trim(Request("hdr"))
ClienteID = Trim(Request("clienteid"))
FacturaID = Trim(Request("facturaid"))
accion = Trim(Request("accion"))

HDR2 = Trim(Request("hdr2"))
ClienteID2 = Trim(Request("clienteid2"))
FacturaID2 = Trim(Request("facturaid2"))

mensajeOk = ""
mensajeError = ""

Dim usuario
usuario = Session("currentUser")

Dim rsFactura, cmd

' Guardar si se envía el formulario
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim motivo
    motivo = Trim(Request.Form("motivo"))
 
    If motivo = "" Then
        mensajeError = "⚠️ Debes ingresar un motivo."
    Else
        Set cmd = Server.CreateObject("ADODB.Command")
        With cmd
            .ActiveConnection = conn
            .CommandText = "[usp_Transportista_HojaDeRuta_Cliente_Justificacion_upd]"
            .CommandType = 4 ' Stored Procedure
            .Parameters.Append .CreateParameter("@HojaDeRutaID", 3, 1, , HDR2)
            .Parameters.Append .CreateParameter("@ClienteID", 3, 1, , ClienteID2)
            .Parameters.Append .CreateParameter("@FacturaID", 3, 1, , FacturaID2)
            .Parameters.Append .CreateParameter("@Accion", 3, 1, , accion)
            .Parameters.Append .CreateParameter("@Motivo", 200, 1, 1000, motivo)
            .Parameters.Append .CreateParameter("@usr", 200, 1, 50, usuario)
            .Execute
        End With
        Set cmd = Nothing
        mensajeOk = "✅ Justificación guardada correctamente."
    End If
End If

' Obtener datos de factura 
sql = "EXEC usp_Transportista_HojaDeRuta_Buscar_Pendientes '2025-01-01'," & _
IIf(HDR = "", "0", HDR) & ", " & IIf(ClienteID2 = "", "0", ClienteID2) & ", " & IIf(FacturaID = "", "0", FacturaID)

Set rsFactura = conn.Execute(sql)

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Justificar Pendientes</title>
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
.boton-container {
    display: flex;
    justify-content: space-between;
    gap: 15px;
    margin-top: 20px;
}

.boton-container button,
.boton-container a.btn-volver {
    flex: 1;
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
.boton-container a.btn-volver:hover {
    background-color: #2c80b4;
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
    <h2>Justificar Motivo</h2>

    <% If mensajeOk <> "" Then %>
            <div class="message-ok">
                <%= mensajeOk %>
            </div>
    <% ElseIf mensajeError <> "" Then %>
        <div class="error"><%= mensajeError %></div>
    <% End If %>

    <% If Not rsFactura.EOF Then %>
        <p><strong>Cliente:</strong> <%= rsFactura("ClienteNombre") %></p>
        <p><strong>Fecha:</strong> <%= FormatDateTime(rsFactura("Fecha"), 2) %></p>
        <p><strong>Importe:</strong> $<%= FormatNumber(rsFactura("ImporteFactura"), 2) %></p>
		<!-- Motivo + botones -->
<form method="post" action="justificar_motivo.asp?hdr2=<%=HDR2%>&hdr=<%=HDR%>&clienteid=<%=ClienteID%>&clienteid2=<%=ClienteID2%>&facturaid=<%=FacturaID%>&facturaid2=<%=FacturaID2%>&accion=<%=accion%>">
    <label for="motivo">Motivo:</label><br/>
    <textarea name="motivo" id="motivo" required></textarea><br/>

    <div class="boton-container">
        <!-- Botón Confirmar -->
        <button type="submit">
            <i class="fas fa-check-circle"></i> Confirmar
        </button>

        <!-- Botón Volver -->
        <button type="button" onclick="window.location.href='validarpendientes.asp?hdr=<%=HDR%>&clienteid=<%=ClienteID%>'">
            <i class="fas fa-arrow-left"></i> Volver
        </button>
    </div>
</form>




    <% Else %>
        <div class="error">⚠️ No se encontró la factura.</div>
    <% End If %>

    <% rsFactura.Close: Set rsFactura = Nothing %>
</div>
</div>

</body>
</html>

<%
Function IfThen(cond, val)
    If cond Then
        IfThen = val
    Else
        IfThen = ""
    End If
End Function

Function IIf(condicion, valorVerdadero, valorFalso)
    If condicion Then
        IIf = valorVerdadero
    Else
        IIf = valorFalso
    End If
End Function
%>
