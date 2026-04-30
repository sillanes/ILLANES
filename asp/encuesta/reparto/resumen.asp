<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("NombreTransportista") = "" Then
    Response.Redirect "/login.asp"
End If

Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor","exec","execute",_
				  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
'Note: We can include following keyword to make a stronger scan but it will also 
'protect users to input these words even those are valid input
'  "!", "char", "alter", "begin", "cast", "create",  
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
 

%><!--#include virtual="./includes/sql-check.asp"--><%

For Each s in Request.Form 
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		' Redirect to an error page
		Response.Redirect(ErrorPage)
	End If
Next
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Transportista</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .diferencia-positiva { color: red; font-weight: bold; }
        .diferencia-negativa { color: green; font-weight: bold; }
        .observacion-box { max-height: 150px; overflow-y: auto; }
    </style>
</head>
<body class="bg-light">
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex:1;">👤 <%=Server.HTMLEncode(Session("NombreTransportista"))%></strong>
    <form method="post" action="logout.asp" style="margin:0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>
<div class="container py-3">
    <h3 class="mb-4 text-center">Resumen Hojas de Ruta</h3>

    <%
    Dim sql, rs
    sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Resumen " & Session("TransportistaID") 
    Set rs = conn.Execute(sql)

    Do While Not rs.EOF
        Dim HojaDeRutaID, Fecha, Gastos, Diferencia, Pendientes, Anuladas, Errores, PendientesJustificadas, Observacion
        HojaDeRutaID = rs("HojaDeRutaID")
        Fecha = rs("Fecha")
        Gastos = rs("Gastos")
        Diferencia = cint(rs("Diferencia"))
        Pendientes = rs("Pendientes")
        Anuladas = rs("Anuladas")
        Errores = rs("Errores")
        PendientesJustificadas = rs("PendientesJustificadas")
        Observacion = rs("Observacion")

        ' Formato diferencia con color
        Dim diferenciaClase
        If Diferencia >= 0 Then
            diferenciaClase = "diferencia-positiva"
        Else
            diferenciaClase = "diferencia-negativa"
        End If
    %>

    <div class="card mb-3 shadow-sm">
        <div class="card-body">
            <h5 class="card-title mb-2">HDR: <%=HojaDeRutaID%></h5>
            <p class="mb-1"><strong>Fecha:</strong> <%=Fecha%></p>
            <p class="mb-1"><strong>Gastos:</strong> $<%=FormatNumber(Gastos, 2)%></p>
            <p class="mb-1"><strong>Diferencia:</strong> <span class="<%=diferenciaClase%>">$<%=FormatNumber(Diferencia, 2)%></span></p>
            <div class="d-flex flex-wrap gap-1 mt-3">
                <span class="badge bg-secondary">📦 Pendientes: <%=Pendientes%></span>
                <span class="badge bg-warning text-dark">❌ Anuladas: <%=Anuladas%></span>
                <span class="badge bg-danger">⚠️ Errores: <%=Errores%></span>
                <span class="badge bg-info text-dark">📝 Pend Injustificadas: <%=PendientesJustificadas%></span>
            </div>
            <button class="btn btn-sm btn-outline-primary mt-3" type="button" data-bs-toggle="collapse" data-bs-target="#obs<%=HojaDeRutaID%>">Ver Observación</button>
            <div class="collapse mt-2" id="obs<%=HojaDeRutaID%>">
                <div class="card card-body observacion-box bg-light">
                    <%=Server.HTMLEncode(Observacion)%>
                </div>
            </div>
        </div>
    </div>

    <%
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    %>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
