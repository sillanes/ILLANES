<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
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


' ==========================================
' Obtener FechaInicio por SP
' ==========================================
Dim cmdFecha, rsFecha, FechaInicioSP
Set cmdFecha = Server.CreateObject("ADODB.Command")
With cmdFecha
    .ActiveConnection = conn
    .CommandType = 4 ' StoredProcedure
    .CommandText = "report.usp_Ordenes_FechaInicio"
End With
Set rsFecha = cmdFecha.Execute

If Not rsFecha.EOF Then
    FechaInicioSP = rsFecha(0)
Else
    FechaInicioSP = Date()
End If
rsFecha.Close
Set rsFecha = Nothing
Set cmdFecha = Nothing

' ==========================================
' Tomar parámetros de fecha del formulario
' ==========================================
Dim FechaInicio, FechaFin
If Request("FechaInicio") <> "" Then
    FechaInicio = CDate(Request("FechaInicio"))
Else
    FechaInicio = FechaInicioSP
End If

If Request("FechaFin") <> "" Then
    FechaFin = CDate(Request("FechaFin"))
Else
    FechaFin = Date()
End If

' ==========================================
' Función para normalizar números
' ==========================================
Function parseNumber(v)
    Dim s, i, dotCount, posLastDot, posLastComma
    If IsNull(v) Then parseNumber = 0 : Exit Function

    ' si es numérico y no contiene coma -> usar directo
    If IsNumeric(v) And InStr(CStr(v), ",") = 0 Then
        On Error Resume Next
        parseNumber = CDbl(v)
        If Err.Number <> 0 Then Err.Clear : parseNumber = 0
        On Error GoTo 0
        Exit Function
    End If

    ' tratar como string
    s = Trim(CStr(v))
    If s = "" Then parseNumber = 0 : Exit Function

    ' limpiar símbolos comunes
    s = Replace(s, "$", "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")

    ' contar puntos
    dotCount = 0
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "." Then dotCount = dotCount + 1
    Next

    posLastDot = InStrRev(s, ".")
    posLastComma = InStrRev(s, ",")

    If InStr(s, ",") > 0 Then
        If InStr(s, ".") > 0 Then
            If posLastComma > posLastDot Then
                ' formato 1.234.567,89
                s = Replace(s, ".", "")
                s = Replace(s, ",", ".")
            Else
                ' formato 1,234,567.89
                s = Replace(s, ",", "")
            End If
        Else
            ' solo coma -> coma decimal
            s = Replace(s, ".", ",")
        End If
    Else
        ' no hay coma
        If dotCount > 1 Then
            s = Replace(s, ",", ".")
        End If
    End If

    On Error Resume Next
    parseNumber = CDbl(s)
    If Err.Number <> 0 Then Err.Clear : parseNumber = 0
    On Error GoTo 0
End Function

' ==========================================
' Función para convertir fecha a ISO yyyy-mm-dd
' ==========================================
Function toISO(fecha)
    Dim d
    d = CDate(fecha)
    toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function

' ==========================================
' Ejecutar SP principal con rango de fechas
' ==========================================
Dim rs, sql
sql = "EXEC report.usp_Ordenes_Ventas_sel '" & toISO(FechaInicio) & "','" & toISO(FechaFin) & "'"
Set rs = conn.Execute(sql)

' ==========================================
' Acumuladores
' ==========================================
Dim totalBronce, totalPlata, totalOro, totalPlatino, totalEspecial, totalImporte
totalBronce = 0
totalPlata = 0
totalOro = 0
totalPlatino = 0
totalEspecial = 0
totalImporte = 0
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Resumen de Ventas</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
    h1 { text-align: center; color: #444; margin-bottom: 20px; }

    .filter-bar { 
        margin-bottom: 20px; 
        display: flex; 
        justify-content: center; 
        gap: 10px; 
        align-items: center; 
        flex-wrap: wrap; /* permite que baje en mobile */
    }
    .filter-bar label { font-weight: bold; }

    .filter-bar input, .filter-bar button { 
        padding: 6px 10px; 
        font-size: 14px; 
    }
    .btn { 
        background-color: #007bff; 
        border: none; 
        color: white; 
        cursor: pointer; 
        border-radius: 4px; 
    }
    .btn:hover { background-color: #0056b3; }


    canvas { width: 100% !important; max-height: 400px; }


    /* -------- RESPONSIVE -------- */
    @media (max-width: 768px) {
        h1 { font-size: 20px; }
        .filter-bar { flex-direction: column; align-items: flex-start; gap: 6px; }
        .filter-bar form { width: 100%; display: flex; flex-direction: column; gap: 8px; }
        .filter-bar input, .filter-bar button { width: 100%; font-size: 15px; }
 
    }
	.progress-container {
    background: #eee;
    border-radius: 10px;
    overflow: hidden;
    height: 18px;
}
.progress-bar {
    height: 100%;
    text-align: center;
    font-size: 11px;
    color: #000;
    font-weight: bold;
    line-height: 18px;
}
.progress-red {
    background: linear-gradient(90deg, #dc3545, #ff7b7b);
}
.progress-yellow {
    background: linear-gradient(90deg, #ffc107, #ffe680);
}
.progress-green {
    background: linear-gradient(90deg, #28a745, #85e085);
}

</style>
</head>
<body>
<!--#include file="header.asp" -->

<div class="main-content">

<h1>Resumen de Ventas</h1>

<!-- Filtro de fechas -->
<div class="filter-bar">
    <form method="get" action="ventas.asp">
        <label>Desde:</label>
        <input type="date" name="FechaInicio" value="<%=toISO(FechaInicio)%>">
        <label>Hasta:</label>
        <input type="date" name="FechaFin" value="<%=toISO(FechaFin)%>">
        <button type="submit" class="btn">Filtrar</button>
    </form>
</div>


<div class="table-container">
<table>
<tr>
    <th>Clientes</th>
    <th>Cajas Oro</th>
    <th>Cajas Plata</th>
    <th>Cajas Platino</th>
    <th>Cajas Especial</th>
    <th>Cajas Bronce</th>
    <th>Importe Total</th>
    <th>Cumplimiento</th>
    <th>Acción</th>
</tr>
<%
Do While Not rs.EOF
    Dim clientes, oro, plata, platino, especial, bronce, importe, faltaCobrar, porcentaje
    clientes = parseNumber(rs("Clientes"))
    oro = parseNumber(rs("CajasOro"))
    plata = parseNumber(rs("CajasPlata"))
    platino = parseNumber(rs("CajasPlatino"))
    especial = parseNumber(rs("CajasEspecial"))
    bronce = parseNumber(rs("CajasBronce"))
    importe = parseNumber(rs("ImporteTotal"))
	
    faltaCobrar  = parseNumber(rs("FaltaCobrar"))
	
	If importe > 0 Then
        porcentaje = ((importe - faltaCobrar) / importe) * 100
    Else
        porcentaje = 0
    End If

    If porcentaje <= 40 Then
        colorClass = "progress-red"
    ElseIf porcentaje <= 80 Then
        colorClass = "progress-yellow"
    Else
        colorClass = "progress-green"
    End If
     
	

    totalOro = totalOro + oro
    totalPlata = totalPlata + plata
    totalPlatino = totalPlatino + platino
    totalEspecial = totalEspecial + especial
    totalBronce = totalBronce + bronce
    totalImporte = totalImporte + importe
%>
<tr>
    <td><%=clientes%></td>
    <td><%=oro%></td>
    <td><%=plata%></td>
    <td><%=platino%></td>
    <td><%=especial%></td>
    <td><%=bronce%></td>
    <td>$<%=FormatNumber(importe,2)%></td>
	<td>
    <div class="progress-container" style="width:120px;">
        <div class="progress-bar <%=colorClass%>" style="width:100%;">
            <%=round(porcentaje,0)%>%
        </div>
    </div>
    </td>
    <td>
        <a href="ventas_detalle.asp?FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>">Ver</a>
    </td>
</tr>
<%
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<tr style="font-weight:bold; background:#f1f1f1;">
    <td>Total</td>
    <td><%=totalOro%></td>
    <td><%=totalPlata%></td>
    <td><%=totalPlatino%></td>
    <td><%=totalEspecial%></td>
    <td><%=totalBronce%></td>
    <td>$<%=FormatNumber(totalImporte,2)%></td>
    <td></td><td></td>
</tr>
</table>

</div>
<canvas id="ventasChart"></canvas>
<script>
const ctx = document.getElementById('ventasChart').getContext('2d');
new Chart(ctx, {
    type: 'bar',
    data: {
        labels: ['Tipos de Cajas'], // un solo grupo
        datasets: [
            {
                label: 'Oro',
                data: [<%=totalOro%>],
                backgroundColor: '#FFD700'
            },
            {
                label: 'Plata',
                data: [<%=totalPlata%>],
                backgroundColor: '#C0C0C0'
            },
            {
                label: 'Platino',
                data: [<%=totalPlatino%>],
                backgroundColor: '#000863'
            },
            {
                label: 'Especial',
                data: [<%=totalEspecial%>],
                backgroundColor: '#FF8C00'
            },
            {
                label: 'Bronce',
                data: [<%=totalBronce%>],
                backgroundColor: '#838739'
            }
        ]
    },
    options: {
        responsive: true,
        plugins: {
            legend: { display: true, position: 'top' },
            datalabels: {
                anchor: 'end',
                align: 'top',
                formatter: function(value) { return value.toLocaleString(); },
                color: '#000',
                font: { weight: 'bold' }
            }
        },
        scales: {
            y: {
                beginAtZero: true,
                ticks: { callback: function(value){ return value.toLocaleString(); } }
            }
        }
    },
    plugins: [ChartDataLabels]
}); 


function toggleSidebar() {
	document.querySelector('.sidebar').classList.toggle('open');
}
</script>


</div>
</body>
</html>
