<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
 
Function FechaLatina(fecha)
    If IsDate(fecha) Then
        FechaLatina = Day(fecha) & "/" & Right("0" & Month(fecha),2) & "/" & Year(fecha)
    Else
        FechaLatina = fecha
    End If
End Function 


If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim vehiculoID
vehiculoID = Trim(Request.QueryString("vehiculoID"))
If vehiculoID = "" Then vehiculoID = "0"
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Mantenimiento Vehículos</title>

<link rel="stylesheet" href="estilos.css" />
<link rel="stylesheet" href="ventas.css" />

<style>
body{
    background:#f4f6f8;
    font-family:Arial,Helvetica,sans-serif;
    margin:0;
}
.container{
    max-width:1000px;
    margin:20px auto;
    padding:0 16px;
}
.card{
    background:#fff;
    border-radius:14px;
    padding:20px;
    margin-bottom:15px;
    box-shadow:0 4px 10px rgba(0,0,0,.08);
}
.title{
    font-size:22px;
    font-weight:800;
    margin-bottom:10px;
    color:#1f2937;
}
select, textarea{
    width:100%;
    padding:10px;
    border-radius:8px;
    border:1px solid #d1d5db;
    font-size:14px;
}

/* Botones */
.btn{
    padding:6px 12px;
    border-radius:6px;
    cursor:pointer;
    border:none;
    font-weight:bold;
    font-size:13px;
}
.btn-blue{ background:#2563eb;color:#fff; }
.btn-red{ background:#dc2626;color:#fff; }
.btn-gray{ background:#6b7280;color:#fff; }

/* Alertas */
.alert-box{
    padding:12px;
    border-radius:10px;
    margin-bottom:8px;
    color:#fff;
    font-weight:600;
}
.alert-verde{ background:#16a34a; }
.alert-amarillo{ background:#f59e0b; }
.alert-rojo{ background:#dc2626; }
.alert-ok{ background:#4b5563; }
.alert-item-title{ font-size:14px; opacity:.9; }
.alert-item-value{ font-size:16px; }

/* Historial */
.timeline-item{
    border-left:3px solid #2563eb;
    padding:8px 12px;
    margin-bottom:10px;
    background:#f9fafb;
    border-radius:6px;
}
.timeline-header{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:8px;
}
.timeline-date{
    color:#1e40af;
    font-weight:700;
    font-size:13px;
}
.timeline-text{
    margin-top:4px;
    white-space:pre-line;
    font-size:14px;
}
</style>

<script>
function cambiarVehiculo(){
    var v = document.getElementById("vehiculo").value;
    if(v === ""){
        window.location = "mantenimiento.asp";
    }else{
        window.location = "mantenimiento.asp?vehiculoID=" + encodeURIComponent(v);
    }
}

/* ====== Historial via AJAX ====== */
function cargarHistorial(){
    var v = document.getElementById("vehiculo").value;
    if(!v){
        document.getElementById("historial").innerHTML = "";
        return;
    }

    fetch("mantenimiento_historial.asp?vehiculoID=" + encodeURIComponent(v))
        .then(r => r.text())
        .then(html => {
            document.getElementById("historial").innerHTML = html;
        });
}

/* ====== Guardar comentario (nuevo) ====== */
function guardarComentario(){
    var v   = document.getElementById("vehiculo").value;
    var txt = document.getElementById("comentario").value.trim();

    if(!v){
        alert("Seleccione un vehículo primero.");
        return;
    }
    if(!txt){
        alert("Debe escribir un comentario.");
        return;
    }

    fetch("mantenimiento_guardar.asp", {
        method:"POST",
        headers:{ "Content-Type":"application/x-www-form-urlencoded" },
        body:"vehiculoID=" + encodeURIComponent(v) +
             "&comentario=" + encodeURIComponent(txt)
    })
    .then(r => r.text())
    .then(() => {
        document.getElementById("comentario").value = "";
        cargarHistorial();
    });
}

/* ====== Editar comentario ====== */
function editar(id){
    var div = document.getElementById("txt_" + id);
    var actual = div.innerText;

    div.setAttribute("data-orig", actual.replace(/'/g,"&#39;"));

    div.innerHTML = ''
        + '<textarea id="edit_' + id + '" style="width:100%;padding:6px;border-radius:6px;border:1px solid #d1d5db;">'
        + actual.replace(/</g,"&lt;").replace(/>/g,"&gt;")
        + '</textarea>'
        + '<div style="margin-top:6px;">'
        + '<button class="btn btn-blue" onclick="guardarEdicion(' + id + ')">Guardar</button> '
        + '<button class="btn btn-gray" onclick="cancelarEdicion(' + id + ')">Cancelar</button>'
        + '</div>';
}

function cancelarEdicion(id){
    var div = document.getElementById("txt_" + id);
    var orig = div.getAttribute("data-orig") || "";
    orig = orig.replace(/&#39;/g,"'");
    div.innerText = orig;
}

function guardarEdicion(id){
    var nuevo = document.getElementById("edit_" + id).value;

    fetch("mantenimiento_editar.asp",{
        method:"POST",
        headers:{ "Content-Type":"application/x-www-form-urlencoded" },
        body:"id=" + encodeURIComponent(id) + "&comentario=" + encodeURIComponent(nuevo)
    })
    .then(r => r.text())
    .then(() => {
        cargarHistorial();
    });
}

/* ====== Eliminar comentario ====== */
function eliminarComentario(id){
    if(!confirm("¿Seguro que desea eliminar este comentario?")) return;

    fetch("mantenimiento_eliminar.asp?id=" + encodeURIComponent(id))
        .then(r => r.text())
        .then(() => {
            cargarHistorial();
        });
}

/* ====== Init ====== */
document.addEventListener("DOMContentLoaded", function(){
    var v = document.getElementById("vehiculo").value;
    if(v){
        cargarHistorial();
    }
});
</script>

</head>
<body>

<!--#include file="header.asp" -->

<div class="container">

    <!-- Selección de vehículo -->
    <div class="card">
        <div class="title">Mantenimiento de Vehículos</div>

        <label><strong>Seleccione vehículo</strong></label>
        <select id="vehiculo" name="vehiculo" onchange="cambiarVehiculo()">
            <option value="">-- Seleccione --</option>
            <%
            Dim rsV, sqlV
            sqlV = "EXEC dbo.usp_Vehiculos_Activos"
            Set rsV = conn.Execute(sqlV)

            Do While Not rsV.EOF
                Dim vid, pat
                vid = rsV("VehiculoID")
                pat = rsV("Patente")

                Response.Write "<option value='" & vid & "'"
                If CStr(vehiculoID) = CStr(vid) Then
                    Response.Write " selected"
                End If
                Response.Write ">" & Server.HTMLEncode(pat) & "</option>"

                rsV.MoveNext
            Loop
            rsV.Close : Set rsV = Nothing
            %>
        </select>
    </div>

<% If vehiculoID <> "0" Then %>

    <!-- Alertas del vehículo -->
    <div class="card">
        <h3 style="margin:0 0 8px;color:#1e40af;">Alertas del vehículo</h3>

        <%
        Dim cmdA, rsA
        Set cmdA = Server.CreateObject("ADODB.Command")
        Set cmdA.ActiveConnection = conn
        cmdA.CommandType = 4
        cmdA.CommandText = "dbo.usp_Vehiculos_Mantenimiento_Alertas"
        cmdA.Parameters.Append cmdA.CreateParameter("@VehiculoID", 3, 1, , CLng(vehiculoID))

        Set rsA = cmdA.Execute()

        Do While Not rsA.EOF
            Dim estado, css
            estado = rsA("Estado")
            Select Case UCase(estado)
                Case "ROJO": css = "alert-rojo"
                Case "AMARILLO": css = "alert-amarillo"
                Case "VERDE": css = "alert-verde"
                Case Else: css = "alert-ok"
            End Select
        %>

            <div class="alert-box <%=css%>">
                <div class="alert-item-title"><%=rsA("Item")%></div>
                <% If Not IsNull(rsA("Valor")) Then %>
                    <div class="alert-item-value"><%=FechaLatina(rsA("Valor"))%></div>
                <% End If %>
                <% If Not IsNull(rsA("Dias")) Then %>
                    <div>Diferencia: <%=rsA("Dias")%> días</div>
                <% End If %>
            </div>

        <%
            rsA.MoveNext
        Loop

        rsA.Close : Set rsA = Nothing
        %>
    </div>

    <!-- Nuevo comentario -->
    <div class="card">
        <h3 style="margin:0 0 8px;color:#1e40af;">Agregar comentario</h3>
        <textarea id="comentario" rows="3" placeholder="Escriba aquí el comentario..."></textarea>
        <button class="btn btn-blue" style="margin-top:8px" onclick="guardarComentario()">Guardar</button>
    </div>

    <!-- Historial -->
    <div class="card">
        <h3 style="margin:0 0 8px;color:#1e40af;">Historial</h3>
        <div id="historial">
            <!-- Se carga por AJAX al iniciar -->
        </div>
    </div>

<% End If %>

</div>

</body>
</html>
