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

Dim API_FIRMADOR_URL
API_FIRMADOR_URL = "http://192.168.200.14:5090/api/firmar-recibo"

Dim RUTA_LOCAL_RECIBOS, RUTA_UNC_RECIBOS, RUTA_LOCAL_FIRMADOS, RUTA_UNC_FIRMADOS
RUTA_LOCAL_RECIBOS  = "C:\RRHH\Archivos\Recibos\"
RUTA_UNC_RECIBOS    = "\\192.168.200.13\RRHH\Archivos\Recibos\"
RUTA_LOCAL_FIRMADOS = "C:\RRHH\Archivos\RecibosFirmados\"
RUTA_UNC_FIRMADOS   = "\\192.168.200.13\RRHH\Archivos\RecibosFirmados\"

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

Function EstadoClase(estado)
    Dim s
    If IsNull(estado) Or IsEmpty(estado) Then
        s = ""
    Else
        s = UCase(Trim(CStr(estado)))
    End If

    Select Case s
        Case "OK", "PROCESADO", "DISPONIBLE", "NOTIFICADO", "LEIDO", "FIRMADO"
            EstadoClase = "ok"
        Case "ERROR"
            EstadoClase = "err"
        Case "PENDIENTE", "PROCESANDO"
            EstadoClase = "pen"
        Case "DUPLICADO"
            EstadoClase = "dup"
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Function PeriodoFormato(p)
    If IsNull(p) Or Trim(CStr(p & "")) = "" Then
        PeriodoFormato = ""
    Else
        Dim a
        a = Split(CStr(p), "-")
        If UBound(a) = 1 Then
            PeriodoFormato = Right("0" & a(1), 2) & "-" & a(0)
        Else
            PeriodoFormato = CStr(p)
        End If
    End If
End Function

Function ObtenerNombreArchivo(path)
    Dim p
    p = InStrRev(path, "\")
    If p > 0 Then
        ObtenerNombreArchivo = Mid(path, p + 1)
    Else
        ObtenerNombreArchivo = path
    End If
End Function

Function ObtenerDirectorioArchivo(path)
    Dim p
    p = InStrRev(path, "\")
    If p > 0 Then
        ObtenerDirectorioArchivo = Left(path, p)
    Else
        ObtenerDirectorioArchivo = ""
    End If
End Function

Function ObtenerExtension(path)
    Dim p
    p = InStrRev(path, ".")
    If p > 0 Then
        ObtenerExtension = Mid(path, p)
    Else
        ObtenerExtension = ""
    End If
End Function

Function GenerarNombreFirmado(nombreOriginal, reciboID)
    Dim base, ext, p
    ext = ObtenerExtension(nombreOriginal)
    If ext = "" Then ext = ".pdf"

    p = InStrRev(nombreOriginal, ".")
    If p > 0 Then
        base = Left(nombreOriginal, p - 1)
    Else
        base = nombreOriginal
    End If

    GenerarNombreFirmado = base & "_firmado_" & reciboID & ext
End Function

Function ReplaceInsensitive(text, findText, replaceText)
    Dim pos, src, look
    src = CStr(text & "")
    look = UCase(src)
    pos = InStr(1, look, UCase(findText), vbTextCompare)

    If pos > 0 Then
        ReplaceInsensitive = Left(src, pos - 1) & replaceText & Mid(src, pos + Len(findText))
    Else
        ReplaceInsensitive = src
    End If
End Function

Function LocalToUNC(pathLocal)
    Dim p
    p = Trim(CStr(pathLocal & ""))

    If Left(p, 2) = "\\" Then
        LocalToUNC = p
        Exit Function
    End If

    If InStr(1, UCase(p), UCase(RUTA_LOCAL_RECIBOS), vbTextCompare) = 1 Then
        LocalToUNC = ReplaceInsensitive(p, RUTA_LOCAL_RECIBOS, RUTA_UNC_RECIBOS)
        Exit Function
    End If

    If InStr(1, UCase(p), UCase(RUTA_LOCAL_FIRMADOS), vbTextCompare) = 1 Then
        LocalToUNC = ReplaceInsensitive(p, RUTA_LOCAL_FIRMADOS, RUTA_UNC_FIRMADOS)
        Exit Function
    End If

    LocalToUNC = p
End Function

Function UNCToLocal(pathUNC)
    Dim p
    p = Trim(CStr(pathUNC & ""))

    If InStr(1, UCase(p), UCase(RUTA_UNC_RECIBOS), vbTextCompare) = 1 Then
        UNCToLocal = ReplaceInsensitive(p, RUTA_UNC_RECIBOS, RUTA_LOCAL_RECIBOS)
        Exit Function
    End If

    If InStr(1, UCase(p), UCase(RUTA_UNC_FIRMADOS), vbTextCompare) = 1 Then
        UNCToLocal = ReplaceInsensitive(p, RUTA_UNC_FIRMADOS, RUTA_LOCAL_FIRMADOS)
        Exit Function
    End If

    UNCToLocal = p
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

Function JsonUnescape(s)
    Dim t
    t = CStr(s & "")
    t = Replace(t, "\\", "\")
    t = Replace(t, "\/", "/")
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\t", vbTab)
    JsonUnescape = t
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

Function HttpPostJson(url, jsonBody, ByRef statusCode, ByRef responseText)
    On Error Resume Next

    Dim http
    statusCode = 0
    responseText = ""

    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 30000, 30000, 30000, 120000
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

Dim empleadoID, reciboid, accion, acepta, nombreFirmante
Dim msgOk, msgErr, salidaApi

empleadoID     = CLng("0" & Session("Empleado_EmpleadoID"))
reciboid       = CLng("0" & Request("reciboid"))
accion         = LCase(Trim(Request("accion")))
acepta         = Trim(Request("acepta"))
nombreFirmante = Trim(Request("nombrefirmante"))

msgOk = ""
msgErr = ""
salidaApi = ""

If reciboid <= 0 Then
    Response.Write "ReciboID inválido."
    Response.End
End If

Dim rs
Set rs = conn.Execute("EXEC empleado.usp_Empleado_Recibo_Detalle_Sel " & empleadoID & ", " & reciboid)

If rs.EOF Then
    Response.Write "No existe el recibo indicado o no pertenece al empleado."
    Response.End
End If

If accion = "firmar" Then
    If acepta <> "1" Then
        msgErr = "Debés confirmar la conformidad para firmar."
    ElseIf nombreFirmante = "" Then
        msgErr = "Debés ingresar tu nombre y apellido."
    ElseIf Trim(Nz(rs("RutaArchivoOriginal"), "")) = "" Then
        msgErr = "No se encontró el PDF original del recibo."
    Else
        Dim rutaOriginalDB, rutaOriginalUNC, nombreOriginal, nombreFirmado
        Dim rutaSalidaUNC, fechaFirma, jsonBody, statusCode, responseText, okApi
        Dim pdfOutput, nombreArchivoOut, hashSha256, apiError, pdfOutputLocal
        Dim dirEntradaUNC, dirRaizRecibosUNC, dirRelativo

        rutaOriginalDB  = Trim(CStr(rs("RutaArchivoOriginal")))
        rutaOriginalUNC = LocalToUNC(rutaOriginalDB)
        nombreOriginal  = ObtenerNombreArchivo(rutaOriginalUNC)
        nombreFirmado   = GenerarNombreFirmado(nombreOriginal, reciboid)

        dirEntradaUNC = ObtenerDirectorioArchivo(rutaOriginalUNC)
        dirRaizRecibosUNC = RUTA_UNC_RECIBOS

        If InStr(1, UCase(dirEntradaUNC), UCase(dirRaizRecibosUNC), vbTextCompare) = 1 Then
            dirRelativo = Mid(dirEntradaUNC, Len(dirRaizRecibosUNC) + 1)
            rutaSalidaUNC = RUTA_UNC_FIRMADOS & dirRelativo & nombreFirmado
        Else
            rutaSalidaUNC = RUTA_UNC_FIRMADOS & nombreFirmado
        End If

        'fechaFirma = Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2) & _
        '             " " & Right("0" & Hour(Now()), 2) & ":" & Right("0" & Minute(Now()), 2)

		fechaFirma = Right("0" & Day(Now()), 2) & "-" & Right("0" & Month(Now()), 2) & "-" & Year(Now()) & _
             " " & Right("0" & Hour(Now()), 2) & ":" & Right("0" & Minute(Now()), 2) 

		jsonBody = "{"
		jsonBody = jsonBody & """pdf_input"":""" & JsSafe(rutaOriginalUNC) & ""","
		jsonBody = jsonBody & """pdf_output"":""" & JsSafe(rutaSalidaUNC) & ""","
		jsonBody = jsonBody & """nombre_firmante"":""" & JsSafe(nombreFirmante) & ""","
		jsonBody = jsonBody & """fecha_firma"":""" & JsSafe(fechaFirma) & ""","
		jsonBody = jsonBody & """page_index"":0,"
		jsonBody = jsonBody & """firma_x"":390,"
		jsonBody = jsonBody & """firma_y"":230,"
		jsonBody = jsonBody & """firma_size"":10,"
		jsonBody = jsonBody & """fecha_x"":390,"
		jsonBody = jsonBody & """fecha_y"":210,"
		jsonBody = jsonBody & """fecha_size"":8"
		jsonBody = jsonBody & "}"

        okApi = HttpPostJson(API_FIRMADOR_URL, jsonBody, statusCode, responseText)
        salidaApi = "POST " & API_FIRMADOR_URL & vbCrLf & _
                    "HTTP " & statusCode & vbCrLf & vbCrLf & responseText

        If Not okApi Then
            msgErr = "No se pudo conectar con el servicio de firmado."
        Else
            If statusCode <> 200 Then
                apiError = JsonUnescape(JsonValue(responseText, "error"))
                If apiError = "" Then apiError = responseText
                msgErr = "La API devolvió error: " & apiError
            Else
                If JsonBool(responseText, "ok") Then
                    pdfOutput        = JsonUnescape(JsonValue(responseText, "pdf_output"))
                    nombreArchivoOut = JsonUnescape(JsonValue(responseText, "nombre_archivo"))
                    hashSha256       = JsonUnescape(JsonValue(responseText, "hash_sha256"))

                    If pdfOutput = "" Or nombreArchivoOut = "" Then
                        msgErr = "La API respondió OK pero no devolvió la ruta del PDF firmado."
                    Else
                        pdfOutputLocal = UNCToLocal(pdfOutput)

                        conn.Execute "EXEC empleado.usp_Empleado_Recibo_Firmar_Archivo_Upd " & _
                                     "@EmpleadoID=" & empleadoID & ", " & _
                                     "@ReciboID=" & reciboid & ", " & _
                                     "@NombreArchivo='" & SqlSafe(nombreArchivoOut) & "', " & _
                                     "@RutaArchivo='" & SqlSafe(pdfOutputLocal) & "', " & _
                                     "@HashArchivo='" & SqlSafe(hashSha256) & "', " & _
                                     "@NombreFirmante='" & SqlSafe(nombreFirmante) & "'"

                        msgOk = "Recibo firmado correctamente y PDF actualizado."

                        rs.Close
                        Set rs = conn.Execute("EXEC empleado.usp_Empleado_Recibo_Detalle_Sel " & empleadoID & ", " & reciboid)
                    End If
                Else
                    apiError = JsonUnescape(JsonValue(responseText, "error"))
                    If apiError = "" Then apiError = responseText
                    msgErr = "No se pudo generar el PDF firmado. " & apiError
                End If
            End If
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Portal del Empleado - Firmar recibo</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css">
<style>
body{margin:0;font-family:Arial, Helvetica, sans-serif;background:#f4f6f9}
.topbar{background:#fff;border-bottom:1px solid #ddd;padding:14px 18px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.topbar .titulo{font-size:22px;font-weight:bold;color:#222}
.topbar .usuario{color:#555;font-size:14px}
.container{max-width:1200px;margin:20px auto;padding:0 12px;box-sizing:border-box}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;background:#fff;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:12px}
.kpi{border:1px solid #e5e5e5;border-radius:10px;padding:12px;background:#fafafa;min-width:0}
.kpi .lab{font-size:12px;color:#666}
.kpi .val{font-size:18px;font-weight:bold;margin-top:4px;word-break:break-word}
.badge{display:inline-block;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:bold}
.badge.ok{background:#d4edda;color:#155724}
.badge.err{background:#f8d7da;color:#721c24}
.badge.pen{background:#fff3cd;color:#856404}
.badge.dup{background:#d1ecf1;color:#0c5460}
.badge.neu{background:#e2e3e5;color:#383d41}
.btn{background:#2b7cff;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.btn:hover{background:#1f68d8}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.btn-sec:hover{background:#5a6268}
.btn-disabled,.btn-disabled:hover{background:#b8c2cc !important;color:#eef2f5 !important;cursor:not-allowed !important}
.alert-ok{background:#d4edda;color:#155724;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.alert-error{background:#f8d7da;color:#721c24;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.form-row{margin-top:14px}
.form-row label{display:block;font-size:13px;color:#555;margin-bottom:6px}
.form-row input[type=text]{width:100%;padding:12px;border:1px solid #ccc;border-radius:8px;box-sizing:border-box}
.checkbox-row{display:flex;gap:10px;align-items:flex-start;margin-top:14px}
.checkbox-row input[type=checkbox]{margin-top:2px}
.pdf-frame{width:100%;height:75vh;border:1px solid #ddd;border-radius:10px;background:#fff}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}
.nota{font-size:12px;color:#666;margin-top:10px;line-height:1.4}
pre{margin:0;white-space:pre-wrap;word-break:break-word;font-size:12px;background:#111;color:#ddd;padding:14px;border-radius:10px;overflow:auto}
@media (max-width: 900px){.grid{grid-template-columns:repeat(2,minmax(160px,1fr))}}
@media (max-width: 768px){
    .container{margin:12px auto;padding:0 10px}
    .card{padding:14px}
    .grid{grid-template-columns:1fr}
    .kpi .val{font-size:16px}
    .actions a,.actions button{width:100%;text-align:center;box-sizing:border-box}
    .topbar{padding:12px}
    .topbar .titulo{font-size:20px}
    .pdf-frame{height:60vh}
}
</style>
<script>
function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}

function toggleFirmaButton() {
    var chk = document.getElementById('acepta');
    var btn = document.getElementById('btnFirmar');
    if (!chk || !btn) return;

    if (chk.checked) {
        btn.disabled = false;
        btn.classList.remove('btn-disabled');
    } else {
        btn.disabled = true;
        btn.classList.add('btn-disabled');
    }
}
</script>
</head>
<body>


<!--#include file="header_empleado.asp" --> 
 

<div class="main-content">
    <% If msgOk <> "" Then %>
        <div class="alert-ok"><%=Server.HTMLEncode(msgOk)%></div>
    <% End If %>

    <% If msgErr <> "" Then %>
        <div class="alert-error"><%=Server.HTMLEncode(msgErr)%></div>
    <% End If %>

    <div class="card">
        <div class="grid">
            <div class="kpi">
                <div class="lab">ReciboID</div>
                <div class="val"><%=Nz(rs("ReciboID"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Período</div>
                <div class="val"><%=PeriodoFormato(rs("Periodo"))%></div>
            </div>
            <div class="kpi">
                <div class="lab">Estado</div>
                <div class="val">
                    <span class="badge <%=EstadoClase(Nz(rs("Estado"), ""))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                </div>
            </div>
        </div>
    </div>


	<div class="card">
		<h3 style="margin-top:0;">Vista previa del PDF</h3>

		<% Dim tieneFirmado
		   tieneFirmado = Trim(Nz(rs("RutaArchivoFirmado"), "")) <> "" %>

		<% If tieneFirmado Then %>
			<iframe class="pdf-frame" src="empleado_recibo_pdf.asp?reciboid=<%=reciboid%>&modo=firmado"></iframe>
		<% Else %>
			<iframe class="pdf-frame" src="empleado_recibo_pdf.asp?reciboid=<%=reciboid%>&modo=original"></iframe>
		<% End If %>

		<div class="actions">
			<% If tieneFirmado Then %>

				<a href="empleado_recibo_pdf.asp?reciboid=<%=reciboid%>&modo=firmado" target="_blank" class="btn">
					📄 Ver PDF firmado
				</a>

				<a href="empleado_recibo_pdf.asp?reciboid=<%=reciboid%>&modo=original" target="_blank" class="btn-sec">
					📄 Ver PDF original
				</a>

			<% Else %>

				<a href="empleado_recibo_pdf.asp?reciboid=<%=reciboid%>&modo=original" target="_blank" class="btn-sec">
					📄 Ver PDF original
				</a>

			<% End If %>
		</div>

		<div class="nota">
			<% If tieneFirmado Then %>
				Estás visualizando el documento firmado. Podés consultar también el original si lo necesitás.
			<% Else %>
				Estás visualizando el documento original. Al firmar, se generará una nueva versión firmada.
			<% End If %>
		</div>
	</div>


    <div class="card">
        <h3 style="margin-top:0;">Confirmación de firma</h3>

        <form method="post" action="empleado_recibo_firmar.asp">
            <input type="hidden" name="reciboid" value="<%=reciboid%>">
            <input type="hidden" name="accion" value="firmar">

            <div class="form-row">
                <label for="nombrefirmante">Nombre y apellido del firmante</label>
                <input type="text" name="nombrefirmante" id="nombrefirmante" value="<%=Server.HTMLEncode(Nz(Session("Empleado_Nombre"), ""))%>">
            </div>

            <div class="checkbox-row">
                <input type="checkbox" name="acepta" id="acepta" value="1" onclick="toggleFirmaButton();">
                <label for="acepta" style="margin:0;">
                    Confirmo que visualicé el recibo, acepto la inserción de mi nombre, apellido, fecha y hora de firma en el documento y firmo su recepción/conformidad desde el portal.
                </label>
            </div>

            <div class="actions">
                <button type="submit" id="btnFirmar" class="btn btn-disabled" disabled>Firmar recibo</button>
            </div>
        </form>
    </div>

<!--   <% If salidaApi <> "" Then %>  -->
<!--   <div class="card"> -->
<!--       <h3 style="margin-top:0;">Salida API</h3>-->
<!--       <pre><%=Server.HTMLEncode(salidaApi)%></pre>-->
<!--   </div>-->
<!--   <% End If %>-->
</div>

<script>
toggleFirmaButton();
</script>

<%
If IsObject(rs) Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>