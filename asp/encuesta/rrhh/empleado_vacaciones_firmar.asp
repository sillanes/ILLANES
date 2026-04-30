<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar_empleados.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 3600

If CLng("0" & Session("Empleado_UsuarioID")) <= 0 Or CLng("0" & Session("Empleado_EmpleadoID")) <= 0 Then
    Response.Redirect "empleado_login.asp"
End If
Const TEMPLATE_PATH = "\\192.168.200.13\RRHH\Plantillas\Vacaciones\PlantillaVacaciones.docx"
Const OUTPUT_DIR    = "\\192.168.200.13\RRHH\Archivos\Vacaciones\"
Const API_DOC_URL   = "http://192.168.200.14:5090/api/vacaciones/generar-y-firmar"

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function SqlSafe(v)
    SqlSafe = Replace(Trim(CStr(v & "")), "'", "''")
End Function

Function JsSafe(v)
    Dim s
    s = CStr(v & "")
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsSafe = s
End Function

Function FechaHoraActualISO()
    FechaHoraActualISO = Year(Now()) & "-" & _
                         Right("0" & Month(Now()), 2) & "-" & _
                         Right("0" & Day(Now()), 2) & " " & _
                         Right("0" & Hour(Now()), 2) & ":" & _
                         Right("0" & Minute(Now()), 2)
End Function

Function HttpPostJson(url, jsonBody, ByRef statusCode, ByRef responseText)
    On Error Resume Next

    Dim http
    statusCode = 0
    responseText = ""

    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 30000, 30000, 30000, 300000
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send jsonBody

    If Err.Number <> 0 Then
        HttpPostJson = False
        responseText = "ERROR HTTP: " & Err.Description
        Err.Clear
        Set http = Nothing
        Exit Function
    End If

    statusCode = CLng(0 & http.Status)
    responseText = http.responseText

    HttpPostJson = True
    Set http = Nothing
    On Error GoTo 0
End Function

Function JsonValue(json, key)
    Dim re, matches
    JsonValue = ""

    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = False
    re.MultiLine = True
    re.Pattern = """" & key & """" & "\s*:\s*""([^""]*)"""

    If re.Test(json) Then
        Set matches = re.Execute(json)
        JsonValue = matches(0).SubMatches(0)
    End If

    Set re = Nothing
End Function

Function JsonBool(json, key)
    Dim re, matches, v
    JsonBool = False

    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = False
    re.MultiLine = True
    re.Pattern = """" & key & """" & "\s*:\s*(true|false)"

    If re.Test(json) Then
        Set matches = re.Execute(json)
        v = LCase(matches(0).SubMatches(0))
        If v = "true" Then JsonBool = True
    End If

    Set re = Nothing
End Function

Function JsonUnescape(s)
    Dim t
    t = CStr(s & "")
    t = Replace(t, "\\", "\")
    t = Replace(t, """" , Chr(34))
    t = Replace(t, "\/", "/")
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\t", vbTab)
    JsonUnescape = t
End Function

Dim empleadoID, vacacionID, accion
Dim msgErr, salidaApi
Dim rsDoc, sqlDoc

empleadoID  = CLng("0" & Session("Empleado_EmpleadoID"))
vacacionID  = CLng("0" & Request("vacacionid"))
accion      = LCase(Trim(Request("accion")))
msgErr      = ""
salidaApi   = ""

If vacacionID <= 0 Then
    Response.Redirect "empleado_vacaciones.asp?msgerr=" & Server.URLEncode("Vacación inválida.")
End If

sqlDoc = "EXEC rrhh.usp_Vacaciones_Documento_Sel " & _
         "@VacacionID=" & vacacionID & ", " & _
         "@EmpleadoID=" & empleadoID

Set rsDoc = conn.Execute(sqlDoc)

If rsDoc.EOF Then
    rsDoc.Close
    Set rsDoc = Nothing
    Response.Redirect "empleado_vacaciones.asp?msgerr=" & Server.URLEncode("No se encontró la vacación indicada.")
End If

If accion = "firmar" Then
    On Error Resume Next

    Dim nombreFirmante, fechaFirma
    Dim statusCode, responseText, okApi
    Dim pdfOutput, nombreArchivoOut, hashSha256, apiError
    Dim jsonBody

    nombreFirmante = Trim(CStr(Session("Empleado_Nombre") & ""))
    fechaFirma = FechaHoraActualISO()

    jsonBody = "{"
    jsonBody = jsonBody & """template_path"":""" & JsSafe(TEMPLATE_PATH) & ""","
    jsonBody = jsonBody & """output_dir"":""" & JsSafe(OUTPUT_DIR) & ""","
    jsonBody = jsonBody & """vacacion_id"":" & vacacionID & ","
    jsonBody = jsonBody & """nombre_firmante"":""" & JsSafe(nombreFirmante) & ""","
    jsonBody = jsonBody & """fecha_firma"":""" & JsSafe(fechaFirma) & ""","
	jsonBody = jsonBody & """firma_x"":90,"
	jsonBody = jsonBody & """firma_y"":405,"
	jsonBody = jsonBody & """firma_size"":10,"
	jsonBody = jsonBody & """fecha_x"":115,"
	jsonBody = jsonBody & """fecha_y"":385,"
	jsonBody = jsonBody & """fecha_size"":10,"	
    jsonBody = jsonBody & """placeholders"":{"
    jsonBody = jsonBody & """@@NombreYApellido@@"":""" & JsSafe(Nz(rsDoc("NombreYApellido"), "")) & ""","
    jsonBody = jsonBody & """@@LEGAJO@@"":""" & JsSafe(Nz(rsDoc("Legajo"), "")) & ""","
    jsonBody = jsonBody & """@@Desde@@"":""" & JsSafe(Nz(rsDoc("Desde"), "")) & ""","
    jsonBody = jsonBody & """@@Hasta@@"":""" & JsSafe(Nz(rsDoc("Hasta"), "")) & ""","
    jsonBody = jsonBody & """@@HastaFin@@"":""" & JsSafe(Nz(rsDoc("HastaFin"), "")) & ""","
    jsonBody = jsonBody & """@@pendientes@@"":""" & JsSafe(Nz(rsDoc("Pendientes"), "")) & """"
    jsonBody = jsonBody & "}"
    jsonBody = jsonBody & "}"

    okApi = HttpPostJson(API_DOC_URL, jsonBody, statusCode, responseText)
    salidaApi = "POST " & API_DOC_URL & vbCrLf & _
                "HTTP " & statusCode & vbCrLf & vbCrLf & responseText

    If Not okApi Then
        msgErr = "No se pudo conectar con el servicio de generación de documento."
    ElseIf statusCode <> 200 Then
        apiError = JsonValue(responseText, "error")
        If apiError = "" Then apiError = responseText
        msgErr = "La API devolvió error: " & apiError
    ElseIf JsonBool(responseText, "ok") Then

        pdfOutput = JsonUnescape(JsonValue(responseText, "pdf_output"))
        nombreArchivoOut = JsonUnescape(JsonValue(responseText, "nombre_archivo"))
        hashSha256 = JsonUnescape(JsonValue(responseText, "hash_sha256"))

        conn.Execute "EXEC rrhh.usp_Vacaciones_Firmar_Upd " & _
                     "@VacacionID=" & vacacionID & ", " & _
                     "@EmpleadoID=" & empleadoID & ", " & _
                     "@RutaArchivo='" & SqlSafe(pdfOutput) & "', " & _
                     "@NombreArchivo='" & SqlSafe(nombreArchivoOut) & "', " & _
                     "@HashArchivo='" & SqlSafe(hashSha256) & "', " & _
                     "@FirmadoPor='" & SqlSafe(nombreFirmante) & "'"

        If Err.Number <> 0 Then
            msgErr = "Se generó el documento pero no se pudo actualizar la firma: " & Err.Description
            Err.Clear
        Else
            rsDoc.Close
            Set rsDoc = Nothing
            Response.Redirect "empleado_vacaciones.asp?msgok=" & Server.URLEncode("La vacación fue firmada correctamente.")
        End If
    Else
        apiError = JsonValue(responseText, "error")
        If apiError = "" Then apiError = responseText
        msgErr = "No se pudo generar el documento. " & apiError
    End If

    On Error GoTo 0
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Portal del Empleado - Firmar Vacación</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css">
<style>
body{margin:0;font-family:Arial, Helvetica, sans-serif;background:#f4f6f9;}
.topbar{background:#fff;border-bottom:1px solid #ddd;padding:14px 18px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;}
.topbar .titulo{font-size:22px;font-weight:bold;color:#222;}
.container{max-width:900px;margin:20px auto;padding:0 12px;box-sizing:border-box;}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;background:#fff;margin-bottom:20px;}
.grid{display:grid;grid-template-columns:repeat(2,minmax(220px,1fr));gap:14px;}
.info-box{border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fafafa;}
.info-label{display:block;font-size:12px;color:#666;margin-bottom:4px;}
.info-value{font-size:15px;font-weight:bold;color:#222;}
.alert-error{background:#f8d7da;color:#721c24;border-radius:8px;padding:12px 14px;margin-bottom:14px;}
.btn{background:#2b7cff;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer;}
.btn:hover{background:#1f68d8;}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer;}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px;}
.notice{font-size:14px;color:#555;line-height:1.5;margin-top:10px;}
pre{margin:0;white-space:pre-wrap;word-break:break-word;font-size:12px;background:#111;color:#ddd;padding:14px;border-radius:10px;overflow:auto;}
.btn-disabled,.btn-disabled:hover{background:#b8c2cc !important;color:#eef2f5 !important;cursor:not-allowed !important;}

.modal-proceso{
    display:none;
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.45);
    z-index:99999;
    align-items:center;
    justify-content:center;
    padding:20px;
    box-sizing:border-box;
}
.modal-proceso.show{
    display:flex;
}
.modal-proceso-box{
    width:100%;
    max-width:420px;
    background:#fff;
    border-radius:16px;
    padding:26px 22px;
    text-align:center;
    box-shadow:0 20px 50px rgba(0,0,0,.22);
}
.modal-proceso-titulo{
    font-size:20px;
    font-weight:bold;
    color:#222;
    margin:14px 0 8px 0;
}
.modal-proceso-texto{
    font-size:14px;
    color:#666;
    line-height:1.5;
}
.spinner-proceso{
    width:56px;
    height:56px;
    margin:0 auto;
    border:6px solid #dbe7ff;
    border-top:6px solid #2b7cff;
    border-radius:50%;
    animation:giro-proceso 1s linear infinite;
}
body.procesando{
    overflow:hidden;
}
@keyframes giro-proceso{
    from{ transform:rotate(0deg); }
    to{ transform:rotate(360deg); }
}

@media (max-width: 768px){
    .grid{grid-template-columns:1fr}
    .btn,.btn-sec{width:100%;text-align:center;box-sizing:border-box}
}
</style>
<script>
function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}

function enviarFirma() {
    var btn = document.getElementById('btnFirmar');
    var modal = document.getElementById('modalProceso');

    if (btn) {
        btn.disabled = true;
        btn.className = 'btn btn-disabled';
        btn.innerHTML = 'Procesando...';
    }

    if (modal) {
        modal.classList.add('show');
    }

    document.body.classList.add('procesando');
    return true;
}
</script>
</head>
<body>
 
<!--#include file="header_empleado.asp" --> 
 

<div id="modalProceso" class="modal-proceso">
    <div class="modal-proceso-box">
        <div class="spinner-proceso"></div>
        <div class="modal-proceso-titulo">Procesando firma</div>
        <div class="modal-proceso-texto">
            Estamos generando el documento, convirtiéndolo a PDF y registrando tu firma.
            <br><br>
            Por favor aguardá unos segundos.
        </div>
    </div>
</div>

<div class="main-content">

    <% If msgErr <> "" Then %>
        <div class="alert-error"><%=Server.HTMLEncode(msgErr)%></div>
    <% End If %>

    <div class="card">
        <h2 style="margin-top:0;">Resumen</h2>

        <div class="grid">
            <div class="info-box">
                <span class="info-label">Empleado</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rsDoc("NombreYApellido"), ""))%></div>
            </div>

            <div class="info-box">
                <span class="info-label">Legajo</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rsDoc("Legajo"), ""))%></div>
            </div>

            <div class="info-box">
                <span class="info-label">Inicio descanso</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rsDoc("Desde"), ""))%></div>
            </div>

            <div class="info-box">
                <span class="info-label">Finalización descanso</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rsDoc("Hasta"), ""))%></div>
            </div>

            <div class="info-box">
                <span class="info-label">Fecha de reincorporación</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rsDoc("HastaFin"), ""))%></div>
            </div>

            <div class="info-box">
                <span class="info-label">Días pendientes</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rsDoc("Pendientes"), ""))%></div>
            </div>
        </div>

        <div class="notice">
            Al confirmar, se generará el documento desde la plantilla, se convertirá a PDF, se firmará y se registrará la firma en el sistema.
        </div>

        <form method="post" action="empleado_vacaciones_firmar.asp" onsubmit="return enviarFirma();">
            <input type="hidden" name="accion" value="firmar">
            <input type="hidden" name="vacacionid" value="<%=vacacionID%>">

            <div class="actions">
                <button type="submit" class="btn" id="btnFirmar">Confirmar firma</button>
                <a href="empleado_vacaciones.asp" class="btn-sec">Cancelar</a>
            </div>
        </form>
    </div>

    <% If salidaApi <> "" Then %>
    <div class="card">
        <h3 style="margin-top:0;">Salida API</h3>
        <pre><%=Server.HTMLEncode(salidaApi)%></pre>
    </div>
    <% End If %>
</div>

<%
If IsObject(rsDoc) Then
    If rsDoc.State = 1 Then rsDoc.Close
    Set rsDoc = Nothing
End If

If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>