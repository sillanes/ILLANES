<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True
Server.ScriptTimeout = 1200
%>

<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->

<%
If Trim(Session("currentUser") & "") <> "admin" AND Trim(Session("currentUser") & "") <> "rrhh"  Then
    Response.Redirect "../login.asp"
End If

Const adCmdStoredProc = 4
Const adParamInput    = 1
Const adInteger       = 3
Const adVarChar       = 200

Function Nz(v, defaultValue)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Function HtmlEncode(v)
    Dim t
    t = CStr(Nz(v, ""))
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    t = Replace(t, "'", "&#39;")
    HtmlEncode = t
End Function

Function FormatearPeriodo(v)
    Dim s
    s = Trim(CStr(Nz(v, "")))

    If s = "" Then
        FormatearPeriodo = ""
        Exit Function
    End If

    If Len(s) = 7 And Mid(s,5,1) = "-" Then
        FormatearPeriodo = Right(s,2) & "-" & Left(s,4)
        Exit Function
    End If

    FormatearPeriodo = s
End Function

Function TieneCampo(rs, campo)
    On Error Resume Next
    Dim tmp
    tmp = rs(campo)
    If Err.Number <> 0 Then
        TieneCampo = False
        Err.Clear
    Else
        TieneCampo = True
    End If
    On Error GoTo 0
End Function

Function ValorCampo(rs, campo, defecto)
    On Error Resume Next
    Dim tmp
    tmp = rs(campo)
    If Err.Number <> 0 Then
        ValorCampo = defecto
        Err.Clear
    Else
        If IsNull(tmp) Or IsEmpty(tmp) Then
            ValorCampo = defecto
        Else
            ValorCampo = tmp
        End If
    End If
    On Error GoTo 0
End Function

Dim accion, periodo, msgOk, msgErr, usuarioEnvioID
accion         = Trim(Request.Form("accion"))
periodo        = Trim(Request.Form("periodo"))
msgOk          = ""
msgErr         = ""
usuarioEnvioID = Null

If Trim("" & Session("CurrentUser")) <> "" Then
    On Error Resume Next
    usuarioEnvioID = (Session("CurrentUser"))
    If Err.Number <> 0 Then
        usuarioEnvioID = Null
        Err.Clear
    End If
    On Error GoTo 0
End If


If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then

    If accion = "notificar" Then

        If periodo = "" Then
            msgErr = "No se recibió el período a notificar."
        Else
            On Error Resume Next

            Dim cmdNoti, rsNoti
            Dim totalProcesados

            totalProcesados = 0

            Set cmdNoti = Server.CreateObject("ADODB.Command")
            Set cmdNoti.ActiveConnection = conn
            cmdNoti.CommandType = adCmdStoredProc
            cmdNoti.CommandText = "rrhh.usp_Recibos_Notificaciones_Cola_Ins"
            cmdNoti.Parameters.Append cmdNoti.CreateParameter("@Periodo", adVarChar, adParamInput, 7, periodo)

            If IsNull(usuarioEnvioID) Then
                cmdNoti.Parameters.Append cmdNoti.CreateParameter("@UsuarioEnvioID", adVarChar, adParamInput, , Null)
            Else
                cmdNoti.Parameters.Append cmdNoti.CreateParameter("@UsuarioEnvioID", adVarChar, adParamInput, 50 , usuarioEnvioID)
            End If

            Set rsNoti = cmdNoti.Execute

            If Err.Number <> 0 Then
                msgErr = "Error al procesar las notificaciones del período " & HtmlEncode(FormatearPeriodo(periodo)) & ": " & HtmlEncode(Err.Description)
                Err.Clear
            Else
                If IsObject(rsNoti) Then
                    If Not rsNoti.EOF Then
                        If TieneCampo(rsNoti, "TotalProcesados") Then
                            totalProcesados = CLng(ValorCampo(rsNoti, "TotalProcesados", 0))
                        ElseIf TieneCampo(rsNoti, "Cantidad") Then
                            totalProcesados = CLng(ValorCampo(rsNoti, "Cantidad", 0))
                        ElseIf TieneCampo(rsNoti, "CantidadNotificada") Then
                            totalProcesados = CLng(ValorCampo(rsNoti, "CantidadNotificada", 0))
                        End If
                    End If
                End If

                If totalProcesados > 0 Then
                    msgOk = "Se generaron " & totalProcesados & " notificaciones para el período " & HtmlEncode(FormatearPeriodo(periodo)) & "."
                Else
                    msgOk = "Se ejecutó correctamente la notificación del período " & HtmlEncode(FormatearPeriodo(periodo)) & "."
                End If
            End If

            If IsObject(rsNoti) Then
                If rsNoti.State = 1 Then rsNoti.Close
                Set rsNoti = Nothing
            End If

            Set cmdNoti = Nothing
            On Error GoTo 0
        End If

    End If

End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "rrhh.usp_Recibos_Periodos_Notificar_Sel"

On Error Resume Next
Set rs = cmd.Execute
If Err.Number <> 0 Then
    msgErr = "Error al obtener el listado de períodos: " & HtmlEncode(Err.Description)
    Err.Clear
End If
On Error GoTo 0
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Notificaciones Masivas</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
	
    <style>
        * { box-sizing: border-box; }

        body {
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            background: #f4f6f9;
            color: #1f2937;
        }

        .main-content {
            padding: 24px;
        }

        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }

        .page-title {
            margin: 0;
            font-size: 28px;
            font-weight: 700;
            color: #111827;
        }

        .page-subtitle {
            margin: 6px 0 0 0;
            color: #6b7280;
            font-size: 14px;
        }

        .card {
            background: #ffffff;
            border-radius: 16px;
            box-shadow: 0 4px 18px rgba(0,0,0,0.08);
            padding: 20px;
            overflow: hidden;
        }

        .alert {
            border-radius: 12px;
            padding: 14px 16px;
            margin-bottom: 18px;
            font-size: 14px;
        }

        .alert-ok {
            background: #ecfdf5;
            color: #065f46;
            border: 1px solid #a7f3d0;
        }

        .alert-err {
            background: #fef2f2;
            color: #991b1b;
            border: 1px solid #fecaca;
        }

        .table-wrap {
            width: 100%;
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 760px;
        }

        thead th {
            background: #111827;
            color: #fff;
            text-align: left;
            padding: 12px 14px;
            font-size: 13px;
            white-space: nowrap;
        }

        tbody td {
            padding: 12px 14px;
            border-bottom: 1px solid #e5e7eb;
            font-size: 14px;
            vertical-align: middle;
        }

        tbody tr:hover {
            background: #f9fafb;
        }

        .badge-periodo {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: #eff6ff;
            color: #1d4ed8;
            font-weight: 700;
            font-size: 13px;
            white-space: nowrap;
        }

        .btn {
            border: 0;
            border-radius: 10px;
            padding: 10px 14px;
            font-size: 14px;
            font-weight: 700;
            cursor: pointer;
            transition: .2s ease;
            white-space: nowrap;
        }

        .btn-primary {
            background: #2563eb;
            color: #fff;
        }

        .btn-primary:hover {
            background: #1d4ed8;
        }

        .btn-primary:active {
            transform: scale(.98);
        }

        .inline-form {
            display: inline;
            margin: 0;
        }

        .empty-box {
            text-align: center;
            padding: 40px 20px;
            color: #6b7280;
        }

        @media (max-width: 768px) {
            .main-content {
                padding: 14px;
            }

            .page-title {
                font-size: 22px;
            }

            .card {
                padding: 14px;
            }
        }
    </style>

    <script>
        function confirmarNotificacion(periodoMostrado) {
            return confirm("¿Deseás generar las notificaciones del período " + periodoMostrado + "?");
        }
    </script>
</head>
<body>

<!--#include file="header.asp" -->
 


<div class="main-content">

    <div class="page-header">
        <div>
            <h1 class="page-title">Notificaciones Masivas</h1>
            <div class="page-subtitle">Seleccioná un período para generar notificaciones de WhatsApp.</div>
        </div>
    </div>

    <% If msgOk <> "" Then %>
        <div class="alert alert-ok"><%=msgOk%></div>
    <% End If %>

    <% If msgErr <> "" Then %>
        <div class="alert alert-err"><%=msgErr%></div>
    <% End If %>

    <div class="card">
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Período</th>
                        <th>Cantidad de recibos</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                <%
                Dim hayDatos
                hayDatos = False

                If IsObject(rs) Then
                    If Not rs.EOF Then
                        hayDatos = True

                        Do Until rs.EOF
                            Dim periodoBD, cantidadRecibos
                            periodoBD       = ValorCampo(rs, "Periodo", "")
                            cantidadRecibos = ValorCampo(rs, "CantidadRecibos", 0)

                            If periodoBD = "" Then
                                periodoBD = ValorCampo(rs, "PeriodoRecibo", "")
                            End If

                            If cantidadRecibos = 0 Then
                                cantidadRecibos = ValorCampo(rs, "Cantidad", 0)
                            End If
                %>
                    <tr>
                        <td>
                            <span class="badge-periodo"><%=HtmlEncode(FormatearPeriodo(periodoBD))%></span>
                        </td>
                        <td><%=HtmlEncode(cantidadRecibos)%></td>
                        <td>
                            <form method="post" class="inline-form" onsubmit="return confirmarNotificacion('<%=HtmlEncode(FormatearPeriodo(periodoBD))%>');">
                                <input type="hidden" name="accion" value="notificar">
                                <input type="hidden" name="periodo" value="<%=HtmlEncode(periodoBD)%>">
                                <button type="submit" class="btn btn-primary">Notificar</button>
                            </form>
                        </td>
                    </tr>
                <%
                            rs.MoveNext
                        Loop
                    End If
                End If

                If Not hayDatos Then
                %>
                    <tr>
                        <td colspan="3">
                            <div class="empty-box">No hay períodos disponibles para notificar.</div>
                        </td>
                    </tr>
                <%
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>

</div>

</body>
</html>

<%
If IsObject(rs) Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

Set cmd = Nothing

If IsObject(conn) Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>