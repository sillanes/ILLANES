<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim OrdenID,FileID
OrdenID = Request("OrdenID")
FileID = Request("FileID")

If OrdenID = "" Then
    Response.Redirect "ventas_detalle.asp"
End If

Dim FechaInicio, FechaFin
If Request("inicio") <> "" Then
    FechaInicio = CDate(Request("inicio"))
Else
    FechaInicio = DateAdd("m",-1,Date()) ' default 1 mes atrás
End If

If Request("fin") <> "" Then
    FechaFin = CDate(Request("fin"))
Else
    FechaFin = Date()
End If

Function toISO(fecha)
    Dim d
    d = CDate(fecha)
    toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function

' Si se presionó el botón confirmar
If Request("accion") = "confirmar" Then
    Dim cmdEliminar
    Set cmdEliminar = Server.CreateObject("ADODB.Command")
    With cmdEliminar
        .ActiveConnection = conn
        .CommandType = 4 ' StoredProcedure
        .CommandText = "report.usp_Ordenes_Ventas_Eliminar"
        .Parameters.Append .CreateParameter("@FileID", 3, 1, , FileID) ' 3 = adInteger, 1 = adParamInput
        .Parameters.Append .CreateParameter("@OrdenID", 3, 1, , OrdenID) ' 3 = adInteger, 1 = adParamInput
        .Execute
    End With
    Set cmdEliminar = Nothing

    ' Redirigir con mensaje de éxito
    Response.Redirect "ventas_detalle.asp?inicio=" & toISO(FechaInicio) & "&fin= " & toISO(FechaFin) & "&mensaje=Orden+eliminada+correctamente"
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Eliminar Orden</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
    .container { max-width:500px; margin:50px auto; background:#fff; padding:25px; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1); text-align:center; }
    h2 { margin-bottom:20px; color:#444; }
    .btn { display:inline-flex; align-items:center; gap:6px; padding:10px 18px; border-radius:6px; text-decoration:none; font-weight:500; cursor:pointer; margin:5px; }
    .btn-confirm { background:#c0392b; color:#fff; border:none; }
    .btn-confirm:hover { background:#992d22; }
    .btn-cancel { background:#6c757d; color:#fff; border:none; }
    .btn-cancel:hover { background:#545b62; }
</style>
</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
    <div class="container">
        <h2>Eliminar Orden <%=OrdenID%></h2>
        <p>¿Está seguro que desea eliminar esta orden? Esta acción no se puede deshacer.</p>
        <form method="post" action="ventas_eliminar.asp?OrdenID=<%=OrdenID%>&FileID=<%=FileID%>&inicio=<%=toISO(FechaInicio)%>&fin=<%=toISO(FechaFin)%>">
            <input type="hidden" name="accion" value="confirmar">
            <button type="submit" class="btn btn-confirm"><i class="fas fa-trash-alt"></i> Confirmar</button>
            <a href="ventas_detalle.asp?inicio=<%=toISO(FechaInicio)%>&fin=<%=toISO(FechaFin)%>" class="btn btn-cancel"><i class="fas fa-times"></i> Cancelar</a>
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
