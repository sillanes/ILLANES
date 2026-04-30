<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim CampaniaID, Source
CampaniaID = Request.QueryString("campaniaid")
Source = Request.QueryString("source")

If CampaniaID = "" Then
    Response.Write "Falta parámetro CampaniaID"
    Response.End
End If

' --- Procesar acciones de Blacklist ---
If Request("accion") <> "" Then
    Dim numero, accion
    numero = Trim(Request("numero"))
    accion = Trim(Request("accion"))
    
    If accion = "agregar" Then
        conn.Execute "INSERT INTO WhatsAppAPI.sofia.WhatsAPP_API_Blacklist (Destinatario, Comentario) VALUES ('" & numero & "', 'Agregado desde dashboard de campaña')"
    ElseIf accion = "quitar" Then
        conn.Execute "DELETE FROM WhatsAppAPI.sofia.WhatsAPP_API_Blacklist WHERE Destinatario='" & numero & "'"
    End If
End If

' --- Consultar respuestas de la campaña ---
Dim rs, sql
sql = "EXEC report.usp_Campania_WhastApp_Dashboard_Respuestas " & CampaniaID & "," & Source 
Set rs = conn.Execute(sql)
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Respuestas WhatsApp</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
body { font-family: Arial; background:#f4f4f4; margin:0; padding:0; }
.main-content { max-width:1000px; margin:30px auto; padding:20px; background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.1); }
h1 { text-align:center; color:#333; margin-bottom:20px; }
table { width:100%; border-collapse:collapse; }
th, td { padding:10px; text-align:center; border-bottom:1px solid #ddd; }
th { background:#007bff; color:white; }
tr:hover { background:#f8f9fa; }
.btn { padding:6px 12px; border-radius:5px; text-decoration:none; color:white; }
.btn-add { background:#dc3545; }
.btn-add:hover { background:#b02a37; }
.btn-remove { background:#28a745; }
.btn-remove:hover { background:#218838; }
.back { display:inline-block; margin-bottom:15px; text-decoration:none; background:#007bff; color:white; padding:6px 12px; border-radius:5px; }
.back:hover { background:#0056b3; }
.btn2 {
    background: linear-gradient(135deg, #007bff, #0056b3);
    color: #fff; border: none; padding: 10px 20px;
    border-radius: 25px; font-size: 16px; cursor: pointer;
}
.btn2-volver {
    display:inline-flex; align-items:center; gap:6px;
    padding:8px 16px; font-size:14px; color:#fff;
    background:linear-gradient(135deg,#6c757d,#495057);
    border-radius:25px; text-decoration:none;
    font-weight:500; margin-top:20px;
}

</style>
<script>
	function toggleSidebar() {
		document.querySelector('.sidebar').classList.toggle('open');
	}

</script>

</head>

<!--#include file="header.asp" -->
<body>

<div class="main-content">
<a class="btn2" href="campania_dashboard.asp">Volver</a>
<h1>📱 Respuestas de la Campaña #<%=CampaniaID%></h1>

<table>
<tr>
    <th>Teléfono</th> 
    <th>Última Respuesta</th>
    <th>Fecha</th>
    <th>Blacklist</th>
</tr>
<%
Do Until rs.EOF
    Dim btn
    If rs("EnBlacklist") = 1 Then
        btn = "<a class='btn btn-remove' href='campania_respuestas.asp?campaniaid=" & CampaniaID & "&source="& source &"&accion=quitar&numero=" & rs("Numero") & "'>Quitar</a>"
    Else
        btn = "<a class='btn btn-add' href='campania_respuestas.asp?campaniaid=" & CampaniaID & "&source="& source &"&accion=agregar&numero=" & rs("Numero") & "'>Agregar</a>"
    End If

    Response.Write "<tr>"
    Response.Write "<td>" & rs("Numero") & "</td>" 
    Response.Write "<td>" & rs("UltimaRespuesta") & "</td>"
    Response.Write "<td>" & rs("FechaEnvio") & "</td>"
    Response.Write "<td>" & btn & "</td>"
    Response.Write "</tr>"

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
</table>
</div>
</body>
</html>
