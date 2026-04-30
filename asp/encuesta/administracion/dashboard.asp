<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then Response.Redirect "/login.asp"

Function NullIfEmpty(v)
    If Trim(v) = "" Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = v
    End If
End Function
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function
Function RandomColor()
    Randomize
    RandomColor = "rgba(" & Int(Rnd()*255) & "," & Int(Rnd()*255) & "," & Int(Rnd()*255) & ",0.7)"
End Function

' === Filtros recibidos ===
Dim cliente, transportistaid, exportar, mesSeleccionado
Dim fechaini, fechafin
cliente = Request("cliente")
mesSeleccionado = Request("mes")
transportistaid = Request("transportistaid")
exportar = (LCase(Request("exportar")) = "exportar a excel")

If mesSeleccionado <> "" Then
    Dim anio, mes
    anio = CInt(Split(mesSeleccionado, "-")(0))
    mes = CInt(Split(mesSeleccionado, "-")(1))
	
    fechaini = anio & "-" & Right("0" & mes, 2) & "-01"
    fechafin = anio & "-" & Right("0" & mes, 2) & "-30"
	
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <title>Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body { font-family: Arial; margin: 20px; }
        label { margin-right: 10px; }
        .grafico-container { width: 90%; margin: 30px auto; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
    </style>
</head>
<body>

<h2>Dashboard</h2>

<div class="filtros">
    <form method="get" action="dashboard.asp">
        <label>Mes:
            <input type="month" name="mes" value="<%=mesSeleccionado%>">
        </label>

        <label>Transportista:
            <select name="transportistaid">
                <option value="">-- Todos --</option>
                <%
                Set rsT = conn.Execute("EXEC [dbo].[usp_Transportista_sel] 0,0")
                Do Until rsT.EOF
                    Dim tid, tsel
                    tid = rsT("TransportistaID")
                    tsel = ""
                    If transportistaid = CStr(tid) Then tsel = "selected"
                %>
                    <option value="<%=tid%>" <%=tsel%>><%=rsT("Nombre")%></option>
                <%
                    rsT.MoveNext
                Loop
                rsT.Close
                %>
            </select>
        </label>

        <label>Cliente:
            <input type="text" name="cliente" value="<%=cliente%>">
        </label>

        <input type="submit" value="Filtrar">
        <input type="submit" name="exportar" value="Exportar a Excel">
    </form>
</div>

<%
If fechaini <> "" And fechafin <> "" Then
response.write "EXEC [dbo].[usp_Transportista_HojaDeRuta_Dashboard] '" & fechaini & "', '" & fechafin & "'," & NullIfEmpty(transportistaid) & "," & IIf(cliente="","NULL","'" & cliente & "'")
    Dim rs, datos1, datos2, datos3, categorias, key
    Set rs = conn.Execute("EXEC [dbo].[usp_Transportista_HojaDeRuta_Dashboard] '" & fechaini & "', '" & fechafin & "'," & NullIfEmpty(transportistaid) & "," & IIf(cliente="","NULL","'" & cliente & "'") )

    Set datos1 = Server.CreateObject("Scripting.Dictionary")
    Set datos2 = Server.CreateObject("Scripting.Dictionary")
    Set datos3 = Server.CreateObject("Scripting.Dictionary")
    Set categorias = Server.CreateObject("Scripting.Dictionary")

    If exportar Then
        Response.Clear
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader "Content-Disposition", "attachment;filename=dashboard.xls"
        Response.Write "<table border='1'><tr><th>Transportista</th><th>Errores</th><th>Pendientes</th><th>Anuladas</th></tr>"
        Do While Not rs.EOF
            Response.Write "<tr><td>" & rs("Categoria") & "</td><td>" & rs("Valor1") & "</td><td>" & rs("Valor2") & "</td><td>" & rs("Valor3") & "</td></tr>"
            rs.MoveNext
        Loop
        Response.Write "</table>"
        rs.Close
        Set rs = Nothing
        Response.End 
    End If

    If Not rs.EOF Then
%>
    <table>
        <tr><th>Transportista</th><th>Errores</th><th>Pendientes</th><th>Anuladas</th></tr>
<%
        Do While Not rs.EOF
            Dim cat, val1, val2, val3
            cat = rs("Categoria")
            val1 = rs("Valor1")
            val2 = rs("Valor2")
            val3 = rs("Valor3")
            datos1(cat) = val1
            datos2(cat) = val2
            datos3(cat) = val3
            categorias(cat) = ""
%>
            <tr><td><%=cat%></td><td><%=val1%></td><td><%=val2%></td><td><%=val3%></td></tr>
<%
            rs.MoveNext
        Loop
%>
    </table>
<%
    Else
        Response.Write "<p>No hay datos para los filtros seleccionados.</p>"
    End If

    rs.Close
    Set rs = Nothing
End If
%>

<% If Not datos1 Is Nothing And datos1.Count > 0 Then %>
<div class="grafico-container">
    <canvas id="grafico"></canvas>
</div>

<script>
    const labels = [<%
        For Each cat In categorias.Keys
            Response.Write "'" & Replace(cat, "'", "\'") & ","
        Next
    %>];

    const datasets = [
        {
            label: 'Errores',
            data: [<% For Each cat In categorias.Keys : Response.Write datos1(cat) & "," : Next %>],
            backgroundColor: 'rgba(255, 99, 132, 0.7)'
        },
        {
            label: 'Pendientes',
            data: [<% For Each cat In categorias.Keys : Response.Write datos2(cat) & "," : Next %>],
            backgroundColor: 'rgba(255, 159, 64, 0.7)'
        },
        {
            label: 'Anuladas',
            data: [<% For Each cat In categorias.Keys : Response.Write datos3(cat) & "," : Next %>],
            backgroundColor: 'rgba(75, 192, 192, 0.7)'
        }
    ];

    new Chart(document.getElementById('grafico'), {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: 'Errores, Pendientes y Anuladas por Transportista'
                }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true, beginAtZero: true }
            }
        }
    });
</script>
<% End If %>

</body>
</html>

