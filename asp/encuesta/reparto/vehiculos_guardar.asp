<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->

<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("Patente") = "" Then
    Response.Redirect "../login.asp"
End If

Function F(name)
  Dim v: v = Request.Form(name)
  If IsNull(v) Then F = "" Else F = Trim(CStr(v))
End Function

' =============================
'   CAPTURAR PARÁMETROS
' =============================
Dim tipo, fecha, km, vehiculoid, observaciones, dni, detalles_xml, usuario
tipo          = F("tipo")
fecha         = F("fecha")
km            = F("km")
vehiculoid    = F("vehiculoid")
observaciones = F("observaciones")
dni           = F("dni")
detalles_xml  = F("detalles_xml")
usuario       = Session("currentUser")

If tipo="" Or fecha="" Or km="" Or vehiculoid="" Or dni="" Then
    Response.Write "Faltan datos obligatorios."
    Response.End
End If

Dim checklistID
Dim cmd, rs

' =============================
'   GUARDAR ENCABEZADO
' =============================
Dim sqlExec
sqlExec = "EXEC usp_Vehiculos_Checklist_Guardar " & _
          "@VehiculoID=" & CLng(vehiculoid) & "," & _
          "@Tipo='" & Replace(tipo,"'","''") & "'," & _
          "@Fecha='" & Replace(fecha,"'","''") & "'," & _
          "@KmActual=" & CLng(km) & "," & _
          "@Observaciones='" & Replace(observaciones,"'","''") & "'," & _
          "@DNI='" & Replace(dni,"'","''") & "'," & _
          "@Usuario='" & Replace(usuario,"'","''") & "'," & _
          "@DetallesXml=N'" & Replace(detalles_xml,"'","''") & "'"

Set rs = conn.Execute(sqlExec)
  
If Not rs Is Nothing Then
    If Not rs.EOF Then
        checklistID = rs("ChecklistID")
    End If
    rs.Close : Set rs = Nothing
End If

Set cmd = Nothing
conn.Close : Set conn = Nothing

%>


<!--#include file="sidebar.asp" -->
<!--#include file="header.asp" -->

<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Checklist guardado</title>

<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
body{background:#f4f6f8;font-family:Arial;margin:0}
.container{max-width:800px;margin:40px auto;padding:15px}
.card{
    background:white;padding:24px;border-radius:16px;
    box-shadow:0 10px 25px rgba(0,0,0,.08)
}
.ok{
    font-size:20px;font-weight:700;color:#065f46;margin-bottom:10px
}
.btn{
    display:inline-block;margin-top:18px;background:#2563eb;color:white;
    padding:12px 20px;border-radius:12px;font-weight:700;text-decoration:none
}
</style>
</head>

<body>

<div class="container">
  <div class="card">

    <div class="ok">✔ Checklist guardado correctamente</div>

    <div style="margin-top:10px;font-size:16px">
        <strong>ID generado:</strong> <%=checklistID%>
    </div>

    <div style="margin-top:8px;font-size:15px;color:#555">
        Puede volver al inicio para cargar otro control.
    </div>

    <a href="vehiculos_home.asp" class="btn">Volver al inicio</a>

  </div>
</div>

</body>
</html>
