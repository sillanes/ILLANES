<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim OrdenID, FileID
OrdenID = Request("OrdenID")
FileID = Request("FileID")

If OrdenID = "" Then
    Response.Write "Faltan parámetros."
    Response.End
End If

' ===============================
' POST: guardar comisión
' ===============================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim pct, valor, pagado, FechaPago
    pct = CDbl(Replace(Request.Form("pct"), ".", ","))

	Dim v
	v = Trim(Request.Form("valor"))
	If InStr(v, ".") > 0 And InStr(v, ",") = 0 Then
		v = Replace(v, ".", ",")
	End If
	valor = CDbl(v)
	

    pagado = Request.Form("pagado")
    If pagado = "on" Then pagado = 1 Else pagado = 0
    
	FechaPago = Request.Form("fecha_pago")
    If FechaPago = "" Then FechaPago = Null


    On Error Resume Next
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = 4 ' adCmdStoredProc
    cmd.CommandText = "report.usp_Ordenes_Comisiones_Ventas_Registrar"

    cmd.Parameters.Append cmd.CreateParameter("@OrdenID", 3, 1, , OrdenID)
    cmd.Parameters.Append cmd.CreateParameter("@FileID", 3, 1, , FileID)
    cmd.Parameters.Append cmd.CreateParameter("@Porcentaje", 5, 1, , pct)
    cmd.Parameters.Append cmd.CreateParameter("@Valor", 5, 1, , valor)
    cmd.Parameters.Append cmd.CreateParameter("@Pagado", 11, 1, , pagado)
	cmd.Parameters.Append cmd.CreateParameter("@FechaPago", 135, 1, , FechaPago) ' adDBTimeStamp


    cmd.Execute
    Set cmd = Nothing
    On Error GoTo 0

    Response.Redirect "comisiones_ventas.asp?mensaje=" & Server.URLEncode("Comision registrada correctamente")
End If

' ===============================
' Consultar datos de la orden
' ===============================
Dim sql, rs
sql = "EXEC [report].[usp_Ordenes_Ventas_Detalle_sel_ID] " & FileID & ", " & OrdenID
Set rs = conn.Execute(sql)

Dim ClienteID, ClienteNombre, VendedorID, VendedorNombre, Importe, ImporteConIva
Dim ComisionPorcentaje, ComisionImporte, ComisionPagada

ComisionPagada=0

If Not rs.EOF Then
    ClienteID = rs("ClienteID")
    ClienteNombre = rs("Cliente")
    VendedorID = rs("VendedorID")
    VendedorNombre = rs("Vendedor")
    Importe = rs("ImporteSinIva")
    ImporteConIva = rs("Monto")
    On Error Resume Next
    If Not IsNull(rs("ComisionPorcentaje")) Then ComisionPorcentaje = CDbl(rs("ComisionPorcentaje")) Else ComisionPorcentaje = 0
    If Not IsNull(rs("ComisionImporte")) Then ComisionImporte = CDbl(rs("ComisionImporte")) Else ComisionImporte = 0 
	If Not IsNull(rs("ComisionPagada")) Then 
		ComisionPagada = CInt(rs("ComisionPagada")) 
	Else 
		ComisionPagada = 0
	End If

	If Not IsNull(rs("ComisionPagadaFecha")) Then
		FechaPago = rs("ComisionPagadaFecha")
	Else
		FechaPago = ""
	End If
    On Error GoTo 0
Else
    Response.Write "<p>No se encontraron datos para esta orden.</p>"
    Response.End
End If
rs.Close
Set rs = Nothing

' Asegurar valores válidos
If IsNull(ComisionPorcentaje) Or IsEmpty(ComisionPorcentaje) Then ComisionPorcentaje = 0
If IsNull(ComisionImporte) Or IsEmpty(ComisionImporte) Then ComisionImporte = 0
If IsNull(ComisionPagada) Or IsEmpty(ComisionPagada) Then ComisionPagada = 0
Dim FechaPagoISO
If IsDate(FechaPago) Then
    FechaPagoISO = Year(FechaPago) & "-" & Right("0" & Month(FechaPago),2) & "-" & Right("0" & Day(FechaPago),2)
Else
    FechaPagoISO = ""
End If

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Modificar Comisión</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
body {
    font-family: Arial, sans-serif;
    background: #f4f6f9;
    margin: 0;
    padding: 0;
}
.main-content {
    margin-left: 240px;
    padding: 20px;
}
h2 {
    text-align: center;
    margin-bottom: 20px;
    color: #333;
}
.card-container {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 20px;
}
.card {
    background: #fff;
    border-radius: 10px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    padding: 20px;
    width: 280px;
    text-align: center;
}
.card h3 { color: #007bff; }
.section {
    margin-top: 40px;
    background: #fff;
    padding: 25px;
    border-radius: 10px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
}
.radio-group {
    display: flex;
    justify-content: center;
    flex-wrap: wrap;
    gap: 15px;
    margin-top: 15px;
}
.radio-group label {
    background: #f8f9fa;
    border: 1px solid #ccc;
    border-radius: 25px;
    padding: 8px 15px;
    cursor: pointer;
    transition: all 0.3s;
}
.radio-group input[type="radio"] { display: none; }
.radio-group input[type="radio"]:checked + label {
    background: linear-gradient(135deg, #007bff, #0056b3);
    color: #fff;
}
.result {
    text-align: center;
    margin-top: 20px;
    font-size: 22px;
    color: #28a745;
    font-weight: bold;
}
.paid-check {
    text-align: center;
    margin-top: 25px;
    font-size: 16px;
}
.paid-check input { transform: scale(1.4); margin-right: 10px; }

#fechaPagoContainer {
    text-align:center;
    margin-top:15px;
    display:none;
}
#fechaPagoContainer label {
    font-weight:bold;
    margin-right:10px;
}

.btn-container { text-align: center; margin-top: 30px; }
.btn {
    background: linear-gradient(135deg, #007bff, #0056b3);
    color: #fff; border: none; padding: 10px 20px;
    border-radius: 25px; font-size: 16px; cursor: pointer;
}
.btn-volver {
    display:inline-flex; align-items:center; gap:6px;
    padding:8px 16px; font-size:14px; color:#fff;
    background:linear-gradient(135deg,#6c757d,#495057);
    border-radius:25px; text-decoration:none;
    font-weight:500; margin-top:20px;
}
</style>
</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
    <h2>Modificar Comisión</h2>

    <div class="card-container">
        <div class="card">
            <h3>Cliente</h3>
            <p><b><%=ClienteID%></b></p>
            <p><%=ClienteNombre%></p>
        </div>

        <div class="card">
            <h3>Vendedor</h3>
            <p><b><%=VendedorID%></b></p>
            <p><%=VendedorNombre%></p>
        </div>

        <div class="card">
            <h3>Subtotal</h3>
            <p style="font-size:22px; color:#28a745;">$ <%=FormatNumber(Importe,2)%></p>
        </div>

        <div class="card">
            <h3>Importe</h3>
            <p style="font-size:22px; color:#28a745;">$ <%=FormatNumber(ImporteConIva,2)%></p>
        </div>
    </div>

    <div class="section">
        <h3>Calcular Comisión</h3>
        <form method="post">
            <div class="radio-group">
                <input type="radio" id="r2" name="comision" value="2"><label for="r2">2%</label>
                <input type="radio" id="r3" name="comision" value="3"><label for="r3">3%</label>
                <input type="radio" id="r5" name="comision" value="5"><label for="r5">5%</label>
                <input type="radio" id="r6" name="comision" value="6"><label for="r6">6%</label>
                <input type="radio" id="r7" name="comision" value="7"><label for="r7">7%</label>
                <input type="radio" id="r8" name="comision" value="8"><label for="r8">8%</label>
                <input type="radio" id="r9" name="comision" value="9"><label for="r9">9%</label>
            </div>

            <div class="result">
                Comisión: $ <span id="valorComision">0.00</span>
            </div>

            <div class="paid-check">
                <input type="checkbox" id="pagado" name="pagado">
                <label for="pagado">Comisión pagada</label>
            </div>
			
			<div id="fechaPagoContainer">
                <label for="fecha_pago">Fecha de pago:</label> 
				<input type="date" id="fecha_pago" name="fecha_pago" value="<%=FechaPagoISO%>">

            </div>
			
            <input type="hidden" name="pct" id="pct" value="">
            <input type="hidden" name="valor" id="valor" value="">

            <div class="btn-container">
                <button type="submit" class="btn">Guardar Comisión</button><br>
                <a href="comisiones_ventas.asp" class="btn-volver"><i class="fas fa-arrow-left"></i> Volver</a>
            </div>
        </form>
    </div>
</div>

<script>
// ---------- Helpers ----------
function parseVBNumber(v) {
    if (v === null || v === undefined) return 0;
    var s = String(v).trim();
    if (s === "" || s.toLowerCase() === "null" || s.toLowerCase() === "undefined") return 0;
    s = s.replace(/\./g, '').replace(',', '.');
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}
function setLabelValues(selectedRb, valor) {
    document.querySelectorAll('.radio-group input[name="comision"]').forEach(rb => {
        const lbl = rb.nextElementSibling;
        if (!lbl) return;
        if (rb === selectedRb) lbl.textContent = rb.value + "% ($" + Number(valor).toFixed(2) + ")";
        else lbl.textContent = rb.value + "%";
    });
}
function pickImporte(subtotal, conIva) {
    return conIva > 0 ? conIva : subtotal;
}

// ---------- Datos del servidor ----------
const importeSubtotal = parseVBNumber("<%=CStr(Importe)%>");
const importeConIva   = parseVBNumber("<%=CStr(ImporteConIva)%>");
const importe         = pickImporte(importeSubtotal, importeConIva);

const spPct     = parseVBNumber("<%=CStr(ComisionPorcentaje)%>");
let   spImporte = parseVBNumber("<%=CStr(ComisionImporte)%>");
const spPagada  = <%=IIf(CInt(ComisionPagada)=1,"true","false")%>;

// ---------- Elementos ----------
const lblValor = document.getElementById('valorComision');
const hidPct   = document.getElementById('pct');
const hidValor = document.getElementById('valor');
const chkPagado = document.getElementById('pagado');
const fechaPagoContainer = document.getElementById('fechaPagoContainer');

if (chkPagado.checked) fechaPagoContainer.style.display = 'block';
chkPagado.addEventListener('change', ()=> {
    fechaPagoContainer.style.display = chkPagado.checked ? 'block' : 'none';
});

// ---------- Eventos ----------
document.querySelectorAll('input[name="comision"]').forEach(radio => {
    radio.addEventListener('change', e => {
        const pct = parseFloat(e.target.value) || 0;
        const valor = (importe * pct / 100);
        lblValor.textContent = valor.toFixed(2);
        hidPct.value = pct;
        hidValor.value = valor.toFixed(2);
        setLabelValues(e.target, valor);
    });
});

// ---------- Inicialización ----------
(function init(){
    // Si la comisión ya está pagada, marcar el checkbox y mostrar calendario
    if (spPagada) {
        chkPagado.checked = true;
        fechaPagoContainer.style.display = 'block';
    }

    chkPagado.addEventListener('change', ()=> {
        fechaPagoContainer.style.display = chkPagado.checked ? 'block' : 'none';
    });

    // Configurar radios de porcentaje
    if (spPct > 0) {
        const rb = document.querySelector(`input[name="comision"][value="${spPct}"]`);
        if (rb) {
            rb.checked = true;
            const valor = (spImporte > 0 ? spImporte : (importe * spPct / 100));
            lblValor.textContent = Number(valor).toFixed(2);
            hidPct.value = spPct;
            hidValor.value = Number(valor).toFixed(2);
            setLabelValues(rb, valor);
            return;
        }
    }

    const defaultRb = document.querySelector('input[name="comision"][value="5"]');
    if (defaultRb) {
        defaultRb.checked = true;
        const valor = importe * 5 / 100;
        lblValor.textContent = valor.toFixed(2);
        hidPct.value = 5;
        hidValor.value = valor.toFixed(2);
        setLabelValues(defaultRb, valor);
    }
})();

</script>

</body>
</html>

<%
Function IfThen(cond, val)
    If cond Then
        IfThen = val
    Else
        IfThen = ""
    End If
End Function

Function IIf(condicion, valorVerdadero, valorFalso)
    If condicion Then
        IIf = valorVerdadero
    Else
        IIf = valorFalso
    End If
End Function
%>
