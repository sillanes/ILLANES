<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Dim action, altaClienteID, dbCon, cmd, rs, sql

If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="../includes/db_command_const.asp" --><%
%><!--#include file="../includes/db_con_open_ventas.asp" --><%

' Procesar acciones POST
If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    action = Request.Form("action")
    altaClienteID = Request.Form("altaClienteID")

    If action = "procesar" And IsNumeric(altaClienteID) Then
        ' Llamar al SP para procesar la solicitud
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbCon
        cmd.CommandType = adCmdStoredProcedure
        cmd.CommandText = "dbo.usp_Clientes_Alta_Procesar"

        cmd.Parameters.Append cmd.CreateParameter("@AltaClienteID", adParamInt, adParamInput, , CInt(altaClienteID))
        cmd.Parameters.Append cmd.CreateParameter("@UsuarioProcesado", adParamVarchar, adParamInput, 20, Session("currentUser"))

        On Error Resume Next
        cmd.Execute
        If Err.Number <> 0 Then
            Response.Write "<script>alert('Error al procesar la solicitud: " & Replace(Err.Description, "'", "\'") & "');</script>"
            Err.Clear
        Else
            Response.Write "<script>alert('Solicitud procesada correctamente.');</script>"
        End If
        On Error GoTo 0

        Set cmd = Nothing
    End If
End If

' Obtener solicitudes pendientes
Set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT AltaClienteID, VendedorID, VendedorNombre, Nombre, Apellido, " & _
      "Direccion, CUITCUIL, Provincia, Ciudad, ConstanciaAfipArchivo, " & _
      "ConstanciaAfipRuta, FechaSolicitud, UsuarioAlta " & _
      "FROM dbo.Clientes_Alta " & _
      "WHERE Status = 0 " & _
      "ORDER BY FechaSolicitud DESC"

rs.Open sql, dbCon, adOpenStatic, adLockReadOnly

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Administración - Alta de Clientes</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        .content-container {
            margin-left: 250px;
            padding: 20px;
            background-color: #f5f5f5;
            min-height: 100vh;
        }

        header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            background: white;
            padding: 18px 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            margin-bottom: 20px;
            position: sticky;
            top: 0;
            z-index: 10;
        }

        header .header-left {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        header .btn-header {
            border: none;
            border-radius: 6px;
            padding: 10px 16px;
            cursor: pointer;
            color: #ffffff;
            background: #2563eb;
            text-decoration: none;
        }

        header .btn-header.red {
            background: #dc2626;
        }

        .content-header {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }

        .content-title {
            font-size: 24px;
            color: #333;
            margin: 0;
        }

        .content-subtitle {
            color: #666;
            margin: 5px 0 0 0;
        }

        .solicitudes-container {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .solicitud-item {
            border-bottom: 1px solid #e5e7eb;
            padding: 20px;
            transition: background-color 0.2s;
        }

        .solicitud-item:hover {
            background-color: #f9fafb;
        }

        .solicitud-item:last-child {
            border-bottom: none;
        }

        .solicitud-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 15px;
        }

        .solicitud-info h3 {
            margin: 0;
            color: #1f2937;
            font-size: 18px;
        }

        .solicitud-meta {
            color: #6b7280;
            font-size: 14px;
        }

        .solicitud-details {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-bottom: 15px;
        }

        .detail-item {
            display: flex;
            flex-direction: column;
        }

        .detail-label {
            font-weight: bold;
            color: #374151;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }

        .detail-value {
            color: #1f2937;
            font-size: 14px;
        }

        .solicitud-actions {
            display: flex;
            gap: 10px;
            justify-content: flex-end;
        }

        .btn {
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            transition: all 0.2s;
        }

        .btn-primary {
            background: #2563eb;
            color: white;
        }

        .btn-primary:hover {
            background: #1d4ed8;
        }

        .btn-secondary {
            background: #6b7280;
            color: white;
        }

        .btn-secondary:hover {
            background: #4b5563;
        }

        .no-solicitudes {
            text-align: center;
            padding: 40px;
            color: #6b7280;
        }

        .no-solicitudes i {
            font-size: 48px;
            margin-bottom: 15px;
            display: block;
        }

        .file-link {
            color: #2563eb;
            text-decoration: none;
        }

        .file-link:hover {
            text-decoration: underline;
        }

        @media (max-width: 768px) {
            .content-container {
                margin-left: 0;
            }

            .solicitud-details {
                grid-template-columns: 1fr;
            }

            .solicitud-actions {
                flex-direction: column;
            }

            .btn {
                justify-content: center;
            }
        }
    </style>
</head>
<body>
<header>
    <button class="btn-header" onclick="toggleSidebar()"><i class="fas fa-bars"></i> Menú</button>
    <div class="header-left">
        <strong>👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
        <a href="logout.asp" class="btn-header red"><i class="fas fa-sign-out-alt"></i> Cerrar sesión</a>
    </div>
</header>
    <div class="content-container">
        <div class="content-header">
            <h1 class="content-title">Alta de Clientes</h1>
            <p class="content-subtitle">Gestionar solicitudes de alta de nuevos clientes</p>
        </div>

        <div class="solicitudes-container">
            <% If rs.EOF Then %>
                <div class="no-solicitudes">
                    <i class="fas fa-check-circle"></i>
                    <h3>No hay solicitudes pendientes</h3>
                    <p>Todas las solicitudes de alta de clientes han sido procesadas.</p>
                </div>
            <% Else %>
                <% Do While Not rs.EOF %>
                <div class="solicitud-item">
                    <div class="solicitud-header">
                        <div class="solicitud-info">
                            <h3><%=Server.HTMLEncode(rs("Nombre") & " " & rs("Apellido"))%></h3>
                            <div class="solicitud-meta">
                                Solicitud #<%=rs("AltaClienteID")%> •
                                Vendedor: <%=Server.HTMLEncode(rs("VendedorNombre"))%> •
                                Fecha: <%=FormatDateTime(rs("FechaSolicitud"), vbShortDate)%>
                            </div>
                        </div>
                    </div>

                    <div class="solicitud-details">
                        <div class="detail-item">
                            <span class="detail-label">CUIT/CUIL</span>
                            <span class="detail-value"><%=Server.HTMLEncode(rs("CUITCUIL"))%></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">Dirección</span>
                            <span class="detail-value"><%=Server.HTMLEncode(rs("Direccion"))%></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">Provincia</span>
                            <span class="detail-value"><%=Server.HTMLEncode(rs("Provincia"))%></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">Ciudad</span>
                            <span class="detail-value"><%=Server.HTMLEncode(rs("Ciudad"))%></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">Constancia AFIP</span>
                            <span class="detail-value">
                                <a href="<%=Server.HTMLEncode(rs("ConstanciaAfipRuta"))%>" target="_blank" class="file-link">
                                    <i class="fas fa-file-pdf"></i> Ver PDF
                                </a>
                            </span>
                        </div>
                    </div>

                    <div class="solicitud-actions">
                        <form method="post" style="display: inline;">
                            <input type="hidden" name="action" value="procesar">
                            <input type="hidden" name="altaClienteID" value="<%=rs("AltaClienteID")%>">
                            <button type="submit" class="btn btn-primary" onclick="return confirm('¿Está seguro de que desea procesar esta solicitud?')">
                                <i class="fas fa-check"></i> Procesar Solicitud
                            </button>
                        </form>
                    </div>
                </div>
                <%
                rs.MoveNext
                Loop
                %>
            <% End If %>
        </div>
    </div>

    <script>
        // Auto-refresh cada 30 segundos para ver nuevas solicitudes
        setTimeout(function() {
            location.reload();
        }, 30000);
    </script>
</body>
</html>

<%
rs.Close
Set rs = Nothing
dbCon.Close
Set dbCon = Nothing
%>