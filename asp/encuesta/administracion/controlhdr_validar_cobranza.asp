<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim hdrid, clienteid, accion, formapago, motivo, usuario, mensajeError, mensajeOk
mensajeError = ""
mensajeOk = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    hdrid = Trim(Request.Form("hdrid"))
    clienteid = Trim(Request.Form("clienteid"))
    accion = Trim(Request.Form("accion"))
    formapago = Trim(Request.Form("formapago"))
    motivo = Trim(Request.Form("motivo"))

    If accion = "Editar" Then
        Dim nuevaAccion
        nuevaAccion = Trim(Request.Form("nuevaAccion"))
        If nuevaAccion <> "Validar" And nuevaAccion <> "Invalidar" Then
            mensajeError = "Debe seleccionar una acción válida para continuar."
        Else
            accion = nuevaAccion ' Sobrescribimos la acción para que se procese abajo
        End If
    End If
Else
    hdrid = Trim(Request.QueryString("hdrid"))
    clienteid = Trim(Request.QueryString("clienteid"))
    accion = Trim(Request.QueryString("accion"))
    formapago = Trim(Request.QueryString("formapago"))
    motivo = ""
End If

usuario = Session("currentUser")

If hdrid = "" Or clienteid = "" Or accion = "" Then
    Response.Write "<p style='color:red;'>Parámetros insuficientes para continuar.</p>"
    Response.End
End If

' Obtener información del cliente desde el SP
Dim sqlInfo, rsInfo
sqlInfo = "EXEC [dbo].[usp_Transportista_HojaDeRuta_Cabecera_selV2] " & hdrid & ", " & clienteid & ", '" & Replace(formapago, "'", "''") & "'"
Set rsInfo = conn.Execute(sqlInfo)

Dim ClienteNombre, EstadoEntrega, TotalFacturas, ImporteACobrar, TotalCobrado
Dim Efectivo, Cheque, Transferencia, Descuentos, Detalles, Observacion, Reasignado

If Not rsInfo.EOF Then
    ClienteNombre   = rsInfo("ClienteNombre")
    EstadoEntrega   = rsInfo("EstadoEntrega")
    TotalFacturas   = rsInfo("TotalFacturas")
    ImporteACobrar  = rsInfo("ImporteACobrar")
    TotalCobrado    = rsInfo("TotalCobrado")
    Efectivo        = rsInfo("Efectivo")
    Cheque          = rsInfo("Cheque")
    Transferencia   = rsInfo("Transferencia")
    Descuentos      = rsInfo("Descuentos")
    Detalles        = rsInfo("Detalles")
    Observacion     = rsInfo("Observacion")
    Reasignado      = rsInfo("Reasignado")
End If

rsInfo.Close
Set rsInfo = Nothing

Const adInteger = 3
Const adVarChar = 200
Const adParamInput = 1

Dim redireccionar


If (Request.ServerVariables("REQUEST_METHOD") = "POST" And mensajeError = "") or accion = "Validar" Then
	redireccionar = ""
	If motivo = "" and accion = "Validar" Then
		redireccionar = "controlhdr_clientes.asp?hdrid=" & hdrid & "&msgAutoSave=Validación guardada correctamente."
		motivo="O.K."
	End If
    If motivo = "" and accion<>"Validar" Then
        mensajeError = "Debe ingresar un motivo para continuar."
    Else
        Dim sql, cmd
        sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Validacion ?, ?, ?, ?, ?, ?"
        Set cmd = Server.CreateObject("ADODB.Command")
        With cmd
            .ActiveConnection = conn
            .CommandText = sql
            .CommandType = 1
            .Parameters.Append .CreateParameter("@hdrid", adInteger, adParamInput, , CLng(hdrid))
            .Parameters.Append .CreateParameter("@clienteid", adInteger, adParamInput, , clienteid)
            .Parameters.Append .CreateParameter("@accion", adInteger, adParamInput, , IIf(accion = "Validar", 1, 0))
            .Parameters.Append .CreateParameter("@motivo", adVarChar, adParamInput, 1000, motivo)
            .Parameters.Append .CreateParameter("@formapago", adVarChar, adParamInput, 50, formapago)
            .Parameters.Append .CreateParameter("@usuario", adVarChar, adParamInput, 50, usuario)
            .Execute
        End With
        Set cmd = Nothing
		
        mensajeOk = "Validación guardada correctamente."
		
		If redireccionar<>"" Then
			response.redirect redireccionar
		End If 
		
    End If
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Validar Cobranza</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    <style>
        body {
            font-family: sans-serif;
            background-color: #f0f2f5;
            margin: 0;
            padding: 0;
        }
        .content-wrapper {
            padding: 20px;
            max-width: 800px;
            margin: auto;
        }
        .card {
            background-color: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            padding: 20px;
            margin-top: 20px;
        }
        h1 {
            margin-top: 0;
            text-align: center;
        }
        label {
            font-weight: bold;
            display: block;
            margin-top: 15px;
            margin-bottom: 5px;
        }
        textarea, select {
            width: 100%;
            padding: 10px;
            border-radius: 8px;
            border: 1px solid #ccc;
            box-sizing: border-box;
            font-size: 1em;
            resize: vertical;
        }
        textarea { height: 100px; }
        .btn-primary, .btn-secondary {
            padding: 10px 20px;
            border-radius: 8px;
            border: none;
            cursor: pointer;
            font-size: 1em;
            text-decoration: none;
            display: inline-block;
        }
        .btn-primary {
            background-color: #28a745;
            color: white;
        }
        .btn-secondary {
            background-color: #2e30c7;
            color: white;
        }
        .button-group {
            display: flex;
            justify-content: space-between;
            margin-top: 20px;
        }
        .message-error {
            background-color: #f8d7da;
            color: #721c24;
            padding: 12px;
            border-radius: 8px;
            margin-top: 20px;
        }
        .message-ok {
            background-color: #d4edda;
            color: #155724;
            padding: 12px;
            border-radius: 8px;
            margin-top: 20px;
        }
        header {
            display: flex;
            align-items: center;
            background-color: #343a40;
            color: white;
            padding: 10px 20px;
        }
        .menu-toggle {
            background: none;
            border: none;
            color: white;
            font-size: 20px;
            margin-right: 10px;
            cursor: pointer;
        }
        .logout {
            background-color: #dc3545;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 5px;
            cursor: pointer;
        }
        table {
            width: 100%;
            font-size: 0.95em;
            margin-top: 15px;
        }
        td {
            padding: 5px;
            vertical-align: top;
        }
        pre {
            margin: 0;
        }
    </style>
    <script>
        function toggleSidebar() {
            document.querySelector('.sidebar').classList.toggle('open');
        }
    </script>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout" />
    </form>
</header>

<div class="content-wrapper">
    <h1>Validar Cobranza</h1>
    <div class="card">
        <h2 style="margin-top: 0;">Cliente: <%= clienteid %> - <%= ClienteNombre %></h2>

        <table>
            <tr><td><strong>Estado Entrega:</strong></td><td><%= EstadoEntrega %></td></tr>
            <tr><td><strong>Total Facturas:</strong></td><td><%= TotalFacturas %></td></tr>
            <tr><td><strong>Importe a Cobrar:</strong></td><td>$<%= FormatNumber(ImporteACobrar, 2) %></td></tr>
            <tr><td><strong>Total Cobrado:</strong></td><td>$<%= FormatNumber(TotalCobrado, 2) %></td></tr>
            <tr><td><strong>Efectivo:</strong></td><td>$<%= FormatNumber(Efectivo, 2) %></td></tr>
            <tr><td><strong>Cheque:</strong></td><td>$<%= FormatNumber(Cheque, 2) %></td></tr>
            <tr><td><strong>Transferencia:</strong></td><td>$<%= FormatNumber(Transferencia, 2) %></td></tr>
            <tr><td><strong>Descuentos:</strong></td><td>$<%= FormatNumber(Descuentos, 2) %></td></tr>
            <tr><td><strong>Observación/DNI:</strong></td><td><%= Observacion %></td></tr>
            <tr><td><strong>Detalles/Descuentos:</strong></td><td><pre><%= Detalles %></pre></td></tr>
            <tr><td><strong>Reasignado:</strong></td><td><%= IIf(Reasignado, "Sí", "No") %></td></tr>
            <tr><td><strong>Forma Pago:</strong></td><td><%= formapago %></td></tr>
            <tr><td><strong>Acción:</strong></td>
                <% If accion = "Validar" Then %>
                    <td><div class="message-ok"><i class="fas fa-check-circle"></i> <%= accion %></div></td>
                <% ElseIf accion = "Invalidar" Then %>
                    <td><div class="message-error"><i class="fas fa-times-circle"></i> <%= accion %></div></td>
                <% Else %>
                    <td><div style="color: #555;">Editar</div></td>
                <% End If %>
            </tr>
        </table>

        <% If mensajeError <> "" Then %>
            <div class="message-error">
                <i class="fas fa-times-circle"></i> <%= mensajeError %>
            </div>
        <% End If %>

        <% If mensajeOk <> "" Then %>
            <div class="message-ok">
                <i class="fas fa-check-circle"></i> <%= mensajeOk %>
            </div>
            <div class="button-group">
                <a href="controlhdr_clientes.asp?hdrid=<%= hdrid %>" class="btn-secondary">
                    <i class="fas fa-arrow-left"></i> Volver
                </a>
            </div>
        <% ElseIf accion = "Editar" Then %>
            <form method="post" action="controlhdr_validar_cobranza.asp">
                <input type="hidden" name="hdrid" value="<%= hdrid %>" />
                <input type="hidden" name="clienteid" value="<%= clienteid %>" />
                <input type="hidden" name="accion" value="Editar" />
                <input type="hidden" name="formapago" value="<%= formapago %>" />

                <label for="nuevaAccion">Seleccione la nueva acción:</label>
                <select name="nuevaAccion" id="nuevaAccion" required>
                    <option value="">-- Seleccionar --</option>
                    <option value="Validar">✅ Validar</option>
                    <option value="Invalidar">❌ Invalidar</option>
                </select>

                <label for="motivo">Motivo de la validación:</label>
                <textarea name="motivo" id="motivo" required><%= Server.HTMLEncode(motivo) %></textarea>

                <div class="button-group">
                    <button type="submit" class="btn-primary">
                        <i class="fas fa-check"></i> Guardar
                    </button>
                    <a href="controlhdr_clientes.asp?hdrid=<%= hdrid %>" class="btn-secondary">
                        <i class="fas fa-arrow-left"></i> Volver
                    </a>
                </div>
            </form>
        <% Else %>
            <form method="post" action="controlhdr_validar_cobranza.asp">
                <input type="hidden" name="hdrid" value="<%= hdrid %>" />
                <input type="hidden" name="clienteid" value="<%= clienteid %>" />
                <input type="hidden" name="accion" value="<%= accion %>" />
                <input type="hidden" name="formapago" value="<%= formapago %>" />

                <label for="motivo">Motivo de la validación:</label>
                <textarea name="motivo" id="motivo"><%= Server.HTMLEncode(motivo) %></textarea>

                <div class="button-group">
                    <button type="submit" class="btn-primary">
                        <i class="fas fa-check"></i> Guardar
                    </button>
                    <a href="controlhdr_clientes.asp?hdrid=<%= hdrid %>" class="btn-secondary">
                        <i class="fas fa-arrow-left"></i> Volver
                    </a>
                </div>
            </form>
        <% End If %>
    </div>
</div>
</body>
</html>
<%
Function IIf(condicion, valorVerdadero, valorFalso)
    If condicion Then
        IIf = valorVerdadero
    Else
        IIf = valorFalso
    End If
End Function
%>
