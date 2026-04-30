<%@ Language="VBScript" %>

<!--#include file="../empretienda/conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

' -------------------------
' Helpers seguros
' -------------------------
Function SafeCStr(val)
    If IsNull(val) Then
        SafeCStr = ""
    Else 
        SafeCStr = CStr(val)
    End If
End Function

Function FieldExists(rs, fieldName)
    Dim f
    FieldExists = False
    If IsObject(rs) Then
        On Error Resume Next
        For Each f In rs.Fields
            If LCase(f.Name) = LCase(fieldName) Then
                FieldExists = True
                Exit Function
            End If
        Next
        On Error GoTo 0
    End If
End Function

Function GetFieldValue(rs, fieldName)
    If Not FieldExists(rs, fieldName) Then
        GetFieldValue = ""
    Else
        If IsNull(rs(fieldName)) Then
            GetFieldValue = ""
        Else
            GetFieldValue = rs(fieldName)
        End If
    End If
End Function

Function toISO(fecha)
    If Trim(SafeCStr(fecha)) = "" Then
        toISO = ""
    Else
        Dim d : d = CDate(fecha)
        toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
    End If
End Function

' -------------------------
' Variables
' -------------------------
Dim FechaInicio, FechaFin, sql, rs
Dim arrFechasDict, dictTotalesPorTipo, dictTotalClientesPorFecha
Dim labelsArr()
Dim labelCount
Dim itemFechaRaw, itemFecha, itemTipo, itemCant, tmpObs
Dim totalGeneral, dataJson
Dim totalClientesVal, clientesVal
Dim i, k

' -------------------------
' Leer filtro de fechas
' -------------------------
FechaInicio = Request("inicio")
FechaFin = Request("fin")

' -------------------------
' Ejecutar SP
' -------------------------
sql = "EXEC report.usp_Ordenes_Ventas_PorFecha_Tipo "
If Trim(FechaInicio & "") <> "" Then sql = sql & "'" & toISO(FechaInicio) & "'" Else sql = sql & "NULL"
sql = sql & ", "
If Trim(FechaFin & "") <> "" Then sql = sql & "'" & toISO(FechaFin) & "'" Else sql = sql & "NULL"

Set rs = conn.Execute(sql)

' -------------------------
' Inicializar estructuras
' -------------------------
Set arrFechasDict = CreateObject("Scripting.Dictionary")
Set dictTotalesPorTipo = CreateObject("Scripting.Dictionary")
Set dictTotalClientesPorFecha = CreateObject("Scripting.Dictionary")

totalGeneral = 0
ReDim labelsArr(-1) : labelCount = -1

dataJson = "["

Do While Not rs.EOF
    ' Fecha
    itemFechaRaw = rs("FechaEntrega")
    ' Tipo de caja
    itemTipo = SafeCStr(rs("TipoCaja"))
    ' Observaciones
    tmpObs = SafeCStr(rs("Observaciones"))
    ' TotalCajas
    Dim totalCajasVal
    totalCajasVal = rs("TotalCajas")
    If IsNumeric(totalCajasVal) Then
        itemCant = CLng(totalCajasVal)
    Else
        itemCant = 0
    End If
    ' TotalClientes
    totalClientesVal = rs("TotalClientes")
    If IsNumeric(totalClientesVal) Then totalClientesVal = CLng(totalClientesVal) Else totalClientesVal = 0
    ' Clientes
    If IsNull(rs("Clientes")) Then
        clientesVal = ""
    Else
        clientesVal = Trim(rs("Clientes"))
    End If

    ' Normalizar fecha
    If SafeCStr(itemFechaRaw) <> "" And IsDate(itemFechaRaw) Then
        Dim d : d = CDate(itemFechaRaw)
        itemFecha = Right("0" & Day(d),2) & "-" & Right("0" & Month(d),2) & "-" & Year(d)
    ElseIf SafeCStr(itemFechaRaw) <> "" Then
        itemFecha = SafeCStr(itemFechaRaw)
    Else
        itemFecha = "(sin fecha)"
    End If

    ' Crear diccionario por fecha si no existe
    If Not arrFechasDict.Exists(itemFecha) Then
        arrFechasDict.Add itemFecha, CreateObject("Scripting.Dictionary")
        labelCount = labelCount + 1
        ReDim Preserve labelsArr(labelCount)
        labelsArr(labelCount) = itemFecha
        arrFechasDict(itemFecha)("Tipos") = ""
        arrFechasDict(itemFecha)("CantidadTotal") = 0
        arrFechasDict(itemFecha)("TotalClientes") = totalClientesVal
        arrFechasDict(itemFecha)("Clientes") = clientesVal
        arrFechasDict(itemFecha)("Observaciones") = tmpObs
    Else
        ' Concatenar clientes y observaciones
        If clientesVal <> "" Then
            If arrFechasDict(itemFecha)("Clientes") = "" Then
                arrFechasDict(itemFecha)("Clientes") = clientesVal
            Else
                arrFechasDict(itemFecha)("Clientes") = arrFechasDict(itemFecha)("Clientes") & "; " & clientesVal
            End If
        End If
        If tmpObs <> "" Then
            If arrFechasDict(itemFecha)("Observaciones") = "" Then
                arrFechasDict(itemFecha)("Observaciones") = tmpObs
            Else
                arrFechasDict(itemFecha)("Observaciones") = arrFechasDict(itemFecha)("Observaciones") & "; " & tmpObs
            End If
        End If
        If totalClientesVal > arrFechasDict(itemFecha)("TotalClientes") Then
            arrFechasDict(itemFecha)("TotalClientes") = totalClientesVal
        End If
    End If

    ' Concatenar tipos y cantidades
    Dim tipoStr
    tipoStr = itemTipo & ": " & itemCant
    If arrFechasDict(itemFecha)("Tipos") = "" Then
        arrFechasDict(itemFecha)("Tipos") = tipoStr
    Else
        arrFechasDict(itemFecha)("Tipos") = arrFechasDict(itemFecha)("Tipos") & ", " & tipoStr
    End If
    arrFechasDict(itemFecha)("CantidadTotal") = arrFechasDict(itemFecha)("CantidadTotal") + itemCant

    ' Totales por tipo y general
    If Not dictTotalesPorTipo.Exists(itemTipo) Then
        dictTotalesPorTipo.Add itemTipo, itemCant
    Else
        dictTotalesPorTipo(itemTipo) = dictTotalesPorTipo(itemTipo) + itemCant
    End If
    totalGeneral = totalGeneral + itemCant

    ' Total clientes por fecha
    dictTotalClientesPorFecha(itemFecha) = arrFechasDict(itemFecha)("TotalClientes")

    ' JSON para gráfico
    Dim safeObsJS, safeClientesJS
    safeObsJS = Replace(Replace(Server.HTMLEncode(arrFechasDict(itemFecha)("Observaciones")), """", "\""") , vbCrLf, "\n")
    safeClientesJS = Replace(Replace(Server.HTMLEncode(arrFechasDict(itemFecha)("Clientes")), """", "\""") , vbCrLf, "\n")
    dataJson = dataJson & "{""fecha"":""" & Replace(itemFecha,"""","'") & """,""tipos"":""" & Replace(arrFechasDict(itemFecha)("Tipos"),"""","'") & """,""cantidadTotal"":" & arrFechasDict(itemFecha)("CantidadTotal") & ",""obs"":""" & safeObsJS & """,""totalClientes"":" & arrFechasDict(itemFecha)("TotalClientes") & ",""clientes"":""" & safeClientesJS & """},"

    rs.MoveNext
Loop

If Right(dataJson,1) = "," Then dataJson = Left(dataJson, Len(dataJson)-1)
dataJson = dataJson & "]"

rs.Close
Set rs = Nothing
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Cajas por Fecha - Entregas</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
.main-content{ padding:18px; max-width:1200px; margin:0 auto; }
.filter-form { display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:14px; }
.filter-form label{ font-weight:600; margin-right:6px; }
.btn { background:#007bff; color:#fff; border:none; padding:8px 12px; border-radius:6px; cursor:pointer; }
.table-container { width:100%; overflow-x:auto; margin-top:18px; }
table { width:100%; border-collapse:collapse; box-shadow:0 2px 6px rgba(0,0,0,0.06); background:#fff; }
th, td { padding:8px 10px; border:1px solid #e6e6e6; text-align:center; white-space:nowrap; }
th { background:#007bff; color:#fff; }
td:first-child, th:first-child { text-align:left; }
.obs-icon { color:#007bff; text-decoration:none; }
#obsModal { display:none; position:fixed; inset:0; background:rgba(0,0,0,0.45); align-items:center; justify-content:center; z-index:9999; }
#obsModalContent{ background:#fff; padding:18px; border-radius:8px; max-width:680px; width:92%; box-shadow:0 5px 20px rgba(0,0,0,0.2); position:relative; }
#closeModal{ position:absolute; right:14px; top:10px; cursor:pointer; font-size:18px; color:#666; }
.clientes-cell { max-width:320px; white-space:normal; word-break:break-word; text-align:left; padding-left:12px; }
</style>
</head>
<body>
<!--#include file="header.asp" -->

<div class="main-content">
    <h2>📦 Cajas a entregar por fecha</h2>

    <form method="get" class="filter-form">
        <label>Fecha inicio:</label>
        <input type="date" name="inicio" value="<%=toISO(FechaInicio)%>">
        <label>Fecha fin:</label>
        <input type="date" name="fin" value="<%=toISO(FechaFin)%>">
        <button type="submit" class="btn"><i class="fas fa-filter"></i> Filtrar</button>
    </form>

    <div style="max-width:1100px; margin:10px auto;">
        <canvas id="chartCajas" style="width:100%; height:380px;"></canvas>
    </div>

    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th>Fecha Entrega</th>
                    <th>Tipos de Caja</th>
                    <th>Cantidad Total</th>
                    <th>Total Clientes</th>
                    <th>Clientes</th>
                    <th>Observaciones</th>
                </tr>
            </thead>
            <tbody>
            <%
            If labelCount >= 0 Then
                For i = 0 To labelCount
                    itemFecha = labelsArr(i)
                    Dim innerDict
                    Set innerDict = arrFechasDict(itemFecha)
            %>
                <tr>
                    <td><%= Server.HTMLEncode(itemFecha) %></td>
                    <td><%= Server.HTMLEncode(innerDict("Tipos")) %></td>
                    <td style="text-align:right;"><%= innerDict("CantidadTotal") %></td>
                    <td style="text-align:right;"><%= innerDict("TotalClientes") %></td>
                    <td class="clientes-cell"><%= Server.HTMLEncode(innerDict("Clientes")) %></td>
                    <td style="text-align:center;">
                        <%
                        If innerDict("Observaciones") <> "" Then
                        %>
                            <a href="javascript:void(0)" class="obs-icon" onclick="openObsModal('<%=Replace(innerDict("Observaciones"),"""","'")%>')"><i class="fas fa-comment-dots"></i></a>
                        <% Else %>
                            -
                        <% End If %>
                    </td>
                </tr>
            <%
                    Set innerDict = Nothing
                Next
            Else
            %>
                <tr><td colspan="6" style="text-align:center; padding:18px;">No hay registros para el rango seleccionado.</td></tr>
            <%
            End If
            %>
            </tbody>
        </table>
    </div>
</div>

<div id="obsModal">
    <div id="obsModalContent">
        <span id="closeModal">&times;</span>
        <h3>Observaciones</h3>
        <div id="obsText" style="white-space:pre-wrap;"></div>
    </div>
</div>

<script>
const dataCajas = <%=dataJson%>;
const labels = [<% 
For i = 0 To labelCount
    Response.Write "'" & Replace(labelsArr(i), "'", "\'") & "'"
    If i < labelCount Then Response.Write ","
Next
%>];

const tipos = [<%
Dim tipoName, lastKey
lastKey = dictTotalesPorTipo.Count - 1
i = 0
For Each tipoName In dictTotalesPorTipo.Keys
    Response.Write "'" & Replace(tipoName,"'","\'") & "'"
    If i < lastKey Then Response.Write ","
    i = i + 1
Next
%>];

const datasets = tipos.map(function(tipo, idx){
    return {
        label: tipo,
        data: labels.map(function(fecha){
            const entry = dataCajas.find(d => d.fecha === fecha);
            if(entry && entry.tipos.includes(tipo)) {
                const regex = new RegExp(tipo + ": (\\d+)");
                const m = entry.tipos.match(regex);
                return m ? parseInt(m[1],10) : 0;
            }
            return 0;
        }),
        backgroundColor: ["#FFD700","#C0C0C0","#000863","#FF8C00","#838739"][idx % 5]
    };
});

const totalData = labels.map(function(fecha){
    const entry = dataCajas.find(d => d.fecha === fecha);
    return entry ? entry.cantidadTotal : 0;
});

datasets.push({
    label: 'Total General',
    type: 'line',
    data: totalData,
    borderColor: '#111',
    borderWidth: 2,
    pointRadius: 4,
    fill: false,
    order: 2
});

const ctx = document.getElementById('chartCajas').getContext('2d');
new Chart(ctx, {
    type: 'bar',
    data: { labels: labels, datasets: datasets },
    options: {
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        plugins: { legend: { position: 'top' } },
        scales: { x: { stacked: true }, y: { stacked: true, beginAtZero:true } }
    }
});

function openObsModal(text){
    document.getElementById('obsText').innerText = text;
    document.getElementById('obsModal').style.display='flex';
}
document.getElementById('closeModal').addEventListener('click', function(){ document.getElementById('obsModal').style.display='none'; });
window.addEventListener('click', function(e){ if(e.target === document.getElementById('obsModal')) document.getElementById('obsModal').style.display='none'; });
function toggleSidebar(){ document.querySelector('.sidebar').classList.toggle('open'); }
</script>
</body>
</html>
