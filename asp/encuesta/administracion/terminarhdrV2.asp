<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If

Dim hojaRutaID
hojaRutaID = Request.QueryString("hdrid")

If hojaRutaID = "" Or Not IsNumeric(hojaRutaID) Then
    Response.Write "<p style='color:red;'>ID de hoja de ruta inválido.</p>"
    Response.End
End If
Function FormatearImporte(valor)
    Dim limpio, entero, decimalPart, resultado, i, contador

    limpio = Replace(CStr(valor), ".", "") ' Quitar puntos de miles
    limpio = Replace(limpio, ",", ".") ' Cambiar coma por punto decimal

    If IsNumeric(limpio) Then
        valor = CDbl(limpio)
        valor = Round(valor, 2)
        entero = Int(valor)
        decimalPart = Right("00" & CStr(Int((valor - entero) * 100 + 0.5)), 2)

        Dim strEntero
        strEntero = CStr(entero)
        resultado = ""
        contador = 0

        For i = Len(strEntero) To 1 Step -1
            resultado = Mid(strEntero, i, 1) & resultado
            contador = contador + 1
            If (contador Mod 3 = 0) And (i > 1) Then
                resultado = "." & resultado
            End If
        Next

        FormatearImporte = resultado & "," & decimalPart
    Else
        FormatearImporte = "0,00"
    End If
End Function



' Ejecutar stored procedure
Dim sSQL, dbRS
sSQL = "EXEC usp_Transportista_HojaDeRuta_Terminadas_Resumen " & hojaRutaID
Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.Open sSQL, conn

If dbRS.EOF Then
    Response.Write "<p style='color:red;'>No se encontraron datos para la hoja de ruta.</p>"
    dbRS.Close
    Set dbRS = Nothing
    conn.Close
    Set conn = Nothing
    Response.End
End If

' Obtener valores del SP
Dim TotalHDR, Gastos, Cheques, Transferencia, DepositoBancario, Diferencia, Observaciones, Efectivo, Pendientes, Anuladas, Errores, PendientesJustificadas

TotalHDR = dbRS("TotalHDR")
Gastos = dbRS("Gastos")
Cheques = dbRS("Cheques")
Transferencia = dbRS("Transferencias")
DepositoBancario = dbRS("Deposito")
Diferencia = dbRS("Diferencia")
Observaciones = dbRS("Observacion")
Efectivo = dbRS("Efectivo")
Anuladas = dbRS("Anuladas")
Pendientes = dbRS("Pendientes")
Errores = dbRS("Errores")
PendientesJustificadas = dbRS("PendientesJustificadas")

dbRS.Close
Set dbRS = Nothing
conn.Close
Set conn = Nothing
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Resumen Hoja de Ruta</title>
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    <style>
        body {
            margin: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f7fa;
        }
        header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            background-color: #343a40;
            color: white;
            padding: 10px 20px;
        }
        .menu-toggle {
            background: none;
            border: none;
            color: white;
            font-size: 1.5em;
            cursor: pointer;
        }
        .logout {
            background-color: #dc3545;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
        }
        .main-content {
            margin-left: 250px;
            padding: 20px;
            transition: margin-left 0.3s;
        }
        .sidebar.open ~ .main-content {
            margin-left: 0;
        }
        .content-header h2 {
            margin-bottom: 20px;
            font-size: 1.8em;
            color: #333;
        }
        .form-layout {
            display: flex;
            flex-direction: column;
            gap: 20px;
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .form-row {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        .form-group {
            flex: 1;
            min-width: 250px;
            display: flex;
            flex-direction: column;
        }
        .form-group label {
            font-weight: 600;
            margin-bottom: 6px;
            color: #555;
        }
        .input-group {
            display: flex;
            align-items: center;
            border: 1px solid #ccc;
            border-radius: 8px;
            overflow: hidden;
            background-color: white;
        }
        .input-prefix {
            padding: 10px 12px;
            background-color: #f1f1f1;
            font-weight: bold;
            color: #555;
            border-right: 1px solid #ccc;
            font-size: 1em;
        }
        .input-group .form-control {
            border: none;
            padding: 10px;
            flex: 1;
            font-size: 1em;
        }
        .input-group .form-control:focus {
            outline: none;
            box-shadow: none;
        }
        .form-actions {
            display: flex;
            justify-content: flex-end;
        }
        .btn-primary {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 12px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 1em;
            transition: background-color 0.3s;
        }
        .btn-primary:hover {
            background-color: #0056b3;
        }
        @media (max-width: 768px) {
            .form-row {
                flex-direction: column;
            }
        }
    </style>
   <script>
        function toggleSidebar() {
            document.querySelector('.sidebar').classList.toggle('open');
        }

        function formatCurrency(value) {
            value = value.replace(/\./g, '').replace(',', '.');
            if (isNaN(value)) return '';

            let parts = parseFloat(value).toFixed(2).split('.');
            let integerPart = parts[0];
            let decimalPart = parts[1];

            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');

            return integerPart + ',' + decimalPart;
        }

        document.addEventListener('DOMContentLoaded', function() {
            document.querySelectorAll('.currency-input').forEach(input => {
                input.addEventListener('blur', function() {
                    let val = this.value.replace(/[^0-9,\.]/g, '');
                    if (val !== '') {
                        this.value = formatCurrency(val);
                    }
                });

                input.addEventListener('input', function() {
                    let val = this.value.replace(/[^0-9,\.]/g, '');
                    let cleanVal = val.replace(/\./g, '').replace(',', '.');
                    if (!isNaN(cleanVal)) {
                        this.value = val;
                    }
                });
            });
        });
    </script>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content">
    <div class="content-header">
        <h2>Resumen Hoja de Ruta</h2>
    </div>

    <div class="content-body">
        <form action="procesar_terminar_hdr.asp" method="post" class="form-layout">
            <input type="hidden" name="hdrid" value="<%=hojaRutaID%>">

            <!-- Agrupados de a dos con $ -->
            <div class="form-row">
                <div class="form-group">
                    <label>Total HDR:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="TotalHDR" value="<%=FormatearImporte(TotalHDR)%>" class="form-control currency-input">
                    </div>
                </div>
                <div class="form-group">
                    <label>Gastos:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="Gastos" value="<%=FormatearImporte(Gastos)%>" class="form-control currency-input">
                    </div>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label>Efectivo:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="Efectivo" value="<%=FormatearImporte(Efectivo)%>" class="form-control currency-input">
                    </div>
                </div>
                <div class="form-group">
                    <label>Cheques:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="Cheques" value="<%=FormatearImporte(Cheques)%>" class="form-control currency-input">
                    </div>
                </div>
                <div class="form-group">
                    <label>Transferencia:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="Transferencia" value="<%=FormatearImporte(Transferencia)%>" class="form-control currency-input">
                    </div>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label>Deposito Bancario:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="DepositoBancario" value="<%=FormatearImporte(DepositoBancario)%>" class="form-control currency-input">
                    </div>
                </div>
                <div class="form-group">
                    <label>Diferencia:</label>
                    <div class="input-group">
                        <span class="input-prefix">$</span>
                        <input type="text" name="Diferencia" value="<%=FormatearImporte(Diferencia)%>" class="form-control currency-input">
                    </div>
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label>Pendientes:</label>
                    <input type="text" name="Pendientes" value="<%=Pendientes%>" class="form-control">
                </div>
                <div class="form-group">
                    <label>Errores:</label>
                    <input type="text" name="Errores" value="<%=Errores%>" class="form-control">
                </div>
            </div>

            <div class="form-row">
                <div class="form-group">
                    <label>Anuladas:</label>
                    <input type="text" name="Anuladas" value="<%=Anuladas%>" class="form-control">
                </div>
                <div class="form-group">
                    <label>Anuladas Justificadas:</label>
                    <input type="text" name="PendientesJustificadas" value="<%=PendientesJustificadas%>" class="form-control">
                </div>
            </div>

            <div class="form-row">
                <div class="form-group" style="flex:1;">
                    <label>Observaciones:</label>
                    <textarea name="Observaciones" class="form-control" rows="4"><%=Observaciones%></textarea>
                </div>
            </div>
<div class="form-actions-row" style="display: flex; justify-content: space-between; align-items: center; margin-top: 20px;">

    <input type="submit" value="Guardar Cambios" class="btn-primary" style="padding: 10px 20px; border-radius: 8px;">
	
    <a href="controlhdr_clientes.asp?hdrid=<%=hojaRutaID%>" 
       class="icon-btn" 
       title="Volver"
       style="background-color: #2e30c7; color: white; padding: 10px 20px; border-radius: 8px; text-decoration: none; font-size: 1em;">
        <i class="fas fa-arrow-left"></i> Volver
    </a>


</div>
</div>
</body>
</html>
