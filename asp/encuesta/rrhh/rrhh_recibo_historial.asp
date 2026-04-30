<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True

If Trim(Session("currentUser") & "") <> "admin" AND Trim(Session("currentUser") & "") <> "rrhh"  Then
    Response.Redirect "../login.asp"
End If

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function ToInt(v, alt)
    On Error Resume Next
    Dim n
    n = alt
    If Trim(CStr(v & "")) <> "" Then
        n = CLng(v)
        If Err.Number <> 0 Then
            n = alt
            Err.Clear
        End If
    End If
    ToInt = n
    On Error GoTo 0
End Function

Function Html(v)
    Html = Server.HTMLEncode(CStr(Nz(v, "")))
End Function

Function MostrarFecha(v)
    If IsNull(v) Or Trim(CStr(v & "")) = "" Then
        MostrarFecha = ""
    Else
        On Error Resume Next
        MostrarFecha = Right("0" & Day(CDate(v)),2) & "/" & Right("0" & Month(CDate(v)),2) & "/" & Year(CDate(v)) & " " & Right("0" & Hour(CDate(v)),2) & ":" & Right("0" & Minute(CDate(v)),2)
        If Err.Number <> 0 Then
            MostrarFecha = CStr(v)
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

Dim reciboID
reciboID = ToInt(Request.QueryString("reciboid"), 0)

If reciboID <= 0 Then
    Response.Write "Recibo inválido."
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "rrhh.usp_Recibo_Notificaciones_Historial_Sel"
cmd.Parameters.Append cmd.CreateParameter("@ReciboID", 3, 1, , reciboID)

Set rs = cmd.Execute
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Historial de notificaciones</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<script>
function toggleSidebar(){
    var sb = document.querySelector('.sidebar');
    if(sb){ sb.classList.toggle('open'); }
}
</script>
</head>
<body>


<!--#include file="header.asp" -->
 
<div class="main-content">
    <div class="header-row">
        <h2 class="page-title">Historial de notificaciones - Recibo <%=reciboID%></h2>
        <div class="acciones-top">
            <a href="rrhh_notificaciones.asp" class="btn-sec"><i class="fa fa-arrow-left"></i> Volver</a>
        </div>
    </div>

    <div class="card">
        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>Fecha envío</th>
                        <th>Usuario</th>
                        <th>Medio</th>
                        <th>Tipo</th>
                        <th>Estado envío</th>
                        <th>Mensaje</th>
                        <th>Observación</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="7">No hay notificaciones registradas para este recibo.</td>
                    </tr>
                <%
                Else
                    Do Until rs.EOF
                %>
                    <tr>
                        <td><%=Html(MostrarFecha(rs("FechaEnvio")))%></td>
                        <td><%=Html(rs("UsuarioEnvioID"))%></td>
                        <td><%=Html(rs("Medio"))%></td>
                        <td><%=Html(rs("TipoNotificacion"))%></td>
                        <td><%=Html(rs("EstadoEnvio"))%></td>
                        <td><%=Html(rs("Mensaje"))%></td>
                        <td><%=Html(rs("Observacion"))%></td>
                    </tr>
                <%
                        rs.MoveNext
                    Loop
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<%
On Error Resume Next
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
End If
Set rs = Nothing
Set cmd = Nothing
conn.Close
Set conn = Nothing
On Error GoTo 0
%>
</body>
</html>