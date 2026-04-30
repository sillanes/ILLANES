<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="_upload.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 3600

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim CARPETA_UPLOADS, CARPETA_LOGS, BAT_IMPORTADOR
CARPETA_UPLOADS = "C:\Logistica\Facturas"
CARPETA_LOGS    = "C:\Logistica\Logs"
BAT_IMPORTADOR  = "C:\Logistica\Scripts\run_importador.bat"

Dim msg, errMsg, salidaPython, facturaID, archivoSubido, nombreArchivoMostrado
Dim doWhat, step, htmlUpload
msg                   = ""
errMsg                = ""
salidaPython          = ""
facturaID             = ""
archivoSubido         = ""
nombreArchivoMostrado = ""
htmlUpload            = ""

doWhat = CLng("0" & Request("doWhat"))
step   = Trim(Request("step"))

If doWhat = 0 Then
    Session("CAMION_FileUploaded") = ""
    Session("CAMION_FileUploaded_Name") = ""
End If

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else
        IIf = sFalse
    End If
End Function

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function Ref(qs)
    Ref = Request.ServerVariables("SCRIPT_NAME")
    If Trim(qs & "") <> "" Then
        Ref = Ref & "?" & qs
    End If
End Function

Sub CrearCarpetaSiNoExiste(path)
    On Error Resume Next
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    Set fso = Nothing
    On Error GoTo 0
End Sub

Function ObtenerExtension(nombreArchivo)
    Dim p
    p = InStrRev(nombreArchivo, ".")
    If p > 0 Then
        ObtenerExtension = LCase(Mid(nombreArchivo, p + 1))
    Else
        ObtenerExtension = ""
    End If
End Function

Function ObtenerArchivoMasReciente(folderPath)
    On Error Resume Next

    Dim fso, folderObj, fileObj
    Dim ultimoNombre, ultimaFecha

    ultimoNombre = ""
    ultimaFecha = CDate("1900-01-01")

    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        Set folderObj = fso.GetFolder(folderPath)
        For Each fileObj In folderObj.Files
            If LCase(fso.GetExtensionName(fileObj.Name)) = "pdf" Then
                If fileObj.DateLastModified > ultimaFecha Then
                    ultimaFecha = fileObj.DateLastModified
                    ultimoNombre = fileObj.Path
                End If
            End If
        Next
    End If

    ObtenerArchivoMasReciente = ultimoNombre

    Set folderObj = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

Function ExtraerFacturaIDDesdeSalida(salida)
    Dim re, m
    ExtraerFacturaIDDesdeSalida = ""

    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = False
    re.Pattern = "FacturaID\s*=\s*(\d+)"

    If re.Test(salida) Then
        Set m = re.Execute(salida)
        ExtraerFacturaIDDesdeSalida = m(0).SubMatches(0)
    End If

    Set re = Nothing
End Function

Function LeerArchivoTexto(path)
    On Error Resume Next

    Dim fso, ts, contenido
    contenido = ""

    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1, False, 0)
        contenido = ts.ReadAll()
        ts.Close
        Set ts = Nothing
    End If

    Set fso = Nothing
    LeerArchivoTexto = contenido
    On Error GoTo 0
End Function

Sub SleepMs(ms)
    Dim t
    t = Timer
    Do While ((Timer - t) * 1000) < ms
        If Timer < t Then Exit Do
    Loop
End Sub

Function EsperarArchivo(path, segundosMax)
    On Error Resume Next

    Dim fso, i
    EsperarArchivo = False

    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    For i = 1 To segundosMax * 10
        If fso.FileExists(path) Then
            EsperarArchivo = True
            Exit For
        End If
        Call SleepMs(100)
    Next

    Set fso = Nothing
    On Error GoTo 0
End Function

Function GenerarNombreLog()
    Dim s
    Randomize
    s = Year(Now()) & _
        Right("0" & Month(Now()), 2) & _
        Right("0" & Day(Now()), 2) & "_" & _
        Right("0" & Hour(Now()), 2) & _
        Right("0" & Minute(Now()), 2) & _
        Right("0" & Second(Now()), 2) & "_" & _
        CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
    GenerarNombreLog = "importador_camiones_" & s & ".txt"
End Function

Function GenerarPrefijoArchivo()
    Dim s
    Randomize
    s = Year(Now()) & _
        Right("0" & Month(Now()), 2) & _
        Right("0" & Day(Now()), 2) & "_" & _
        Right("0" & Hour(Now()), 2) & _
        Right("0" & Minute(Now()), 2) & _
        Right("0" & Second(Now()), 2) & "_" & _
        CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
    GenerarPrefijoArchivo = s
End Function

Function ProbarEscrituraArchivo(path)
    On Error Resume Next

    Dim fso, ts
    ProbarEscrituraArchivo = ""

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, True)

    If Err.Number <> 0 Then
        ProbarEscrituraArchivo = "ERROR escribiendo archivo: " & Err.Description
        Err.Clear
    Else
        ts.WriteLine "OK"
        ts.Close
        ProbarEscrituraArchivo = "OK"
    End If

    Set ts = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

Function BorrarArchivoSiExiste(path)
    On Error Resume Next

    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) Then
        fso.DeleteFile path, True
    End If

    Set fso = Nothing
    BorrarArchivoSiExiste = True
    On Error GoTo 0
End Function

Function EjecutarComandoRun(cmd)
    On Error Resume Next

    Dim sh, rc
    Set sh = Server.CreateObject("WScript.Shell")
    rc = sh.Run(cmd, 0, True)

    If Err.Number <> 0 Then
        EjecutarComandoRun = "ERROR al ejecutar comando: " & Err.Description & vbCrLf & vbCrLf & "CMD: " & cmd
        Err.Clear
    Else
        EjecutarComandoRun = "ExitCode=" & rc
    End If

    Set sh = Nothing
    On Error GoTo 0
End Function

Function do_Upload_PDF_Camiones()
    Server.ScriptTimeout = 3600

    Dim Form : Set Form = New ASPForm
    Dim upid

    upid = Trim(Request.QueryString("UploadID"))
    If upid <> "" Then Form.UploadID = upid

    Form.SizeLimit = 10 * 1024 * 1024

    On Error Resume Next
    Form.ReadTimeout = 600
    On Error GoTo 0

    Dim HTML, hResult, archivoDetectado, UploadID
    Dim fso, archivoFinal, prefijo

    Const fsCompletted  = 0
    Const fsSizeLimit   = &HD
    Const fsTimeOut     = &HE
    Const fsError       = &HA

    hResult = ""
    archivoDetectado = ""

    If Form.State > fsError Then

        If Form.State = fsSizeLimit Then
            hResult = "El archivo supera el límite permitido (" & Form.SizeLimit/1024 & " KB)."
        ElseIf Form.State = fsTimeOut Then
            hResult = "El tiempo de carga excedió el máximo permitido (" & Form.ReadTimeout & " seg)."
        Else
            hResult = "Error de carga (código " & Form.State & ")."
        End If

        errMsg = hResult
        hResult = "<div class='alert alert-error'>" & hResult & "</div>"
        Response.Status = "400 Bad request"

    ElseIf Form.State = fsCompletted Then

        Call CrearCarpetaSiNoExiste(CARPETA_UPLOADS)

        ' IMPORTANTE: sin barra final
        Form.Files.Save CARPETA_UPLOADS

        archivoDetectado = ObtenerArchivoMasReciente(CARPETA_UPLOADS)

        If Trim(archivoDetectado) = "" Then
            errMsg = "No se encontró el archivo PDF luego de la subida."
            hResult = "<div class='alert alert-error'>" & errMsg & "</div>"
        Else
            If LCase(ObtenerExtension(archivoDetectado)) <> "pdf" Then
                errMsg = "Solo se permiten archivos PDF."
                hResult = "<div class='alert alert-error'>" & errMsg & "</div>"
            Else
                Set fso = Server.CreateObject("Scripting.FileSystemObject")

                prefijo = GenerarPrefijoArchivo()
                archivoFinal = CARPETA_UPLOADS & "\" & prefijo & "_" & fso.GetFileName(archivoDetectado)

                If LCase(archivoDetectado) <> LCase(archivoFinal) Then
                    If fso.FileExists(archivoFinal) Then
                        fso.DeleteFile archivoFinal, True
                    End If
                    fso.MoveFile archivoDetectado, archivoFinal
                End If

                Set fso = Nothing

                Session("CAMION_FileUploaded") = archivoFinal
                Session("CAMION_FileUploaded_Name") = Mid(archivoFinal, InStrRev(archivoFinal, "\") + 1)

                archivoSubido = Session("CAMION_FileUploaded")
                nombreArchivoMostrado = Session("CAMION_FileUploaded_Name")

                hResult = "<div class='alert alert-ok'>Archivo subido correctamente: " & Server.HTMLEncode(nombreArchivoMostrado) & "</div>"
            End If
        End If

    ElseIf LCase(Trim(Request.QueryString("Action"))) = "cancel" Then
        hResult = "<div class='alert alert-error'>La carga fue cancelada.</div>"
    End If

    UploadID = Form.NewUploadID

    HTML = ""
    HTML = HTML & hResult
    HTML = HTML & "<form name='FFUP' method='post' enctype='multipart/form-data' " & _
                  "action='" & Ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&step=1") & "' " & _
                  "onsubmit='return ProgressBarSmart();'>"

    HTML = HTML & "<div class='dropzone' id='dropzone'>"
    HTML = HTML & "  <div><strong>Arrastrá acá el PDF</strong> o hacé click para seleccionarlo</div>"
    HTML = HTML & "  <div style='margin-top:8px;color:#666;'>Solo PDF de facturas</div>"
    HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required accept='application/pdf,.pdf'>"
    HTML = HTML & "  <div class='nombre-archivo' id='nombreArchivo'></div>"
    HTML = HTML & "</div>"

    HTML = HTML & "<div class='acciones'>"
    HTML = HTML & "  <button type='submit' class='btn'>Subir PDF</button>"
    HTML = HTML & "</div>"
    HTML = HTML & "</form>"

    HTML = HTML & "<script>" & vbCrLf
    HTML = HTML & "function isMobile(){ return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent); }" & vbCrLf
    HTML = HTML & "function ProgressBarSmart(){ if(isMobile()){ return true; }" & vbCrLf
    HTML = HTML & "  var ProgressURL='progress.asp?UploadID=" & UploadID & "';" & vbCrLf
    HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
    HTML = HTML & "  return true; }" & vbCrLf
    HTML = HTML & "</script>"

    do_Upload_PDF_Camiones = HTML
End Function

Call CrearCarpetaSiNoExiste(CARPETA_LOGS)
htmlUpload = do_Upload_PDF_Camiones()

If Trim(Session("CAMION_FileUploaded") & "") <> "" Then
    archivoSubido = Session("CAMION_FileUploaded")
    nombreArchivoMostrado = Session("CAMION_FileUploaded_Name")

    If archivoSubido <> "" Then
        Dim cmdExe, cmdFinal, salidaReal, logPath, probePath, nombreLog, salidaCmd, pruebaLog

        nombreLog = GenerarNombreLog()
        logPath   = CARPETA_LOGS & "\" & nombreLog
        probePath = CARPETA_LOGS & "\probe_" & nombreLog

        pruebaLog = ProbarEscrituraArchivo(probePath)

        If pruebaLog <> "OK" Then
            salidaReal = "IIS no puede escribir en la carpeta de logs." & vbCrLf & _
                         pruebaLog & vbCrLf & _
                         "Ruta: " & CARPETA_LOGS

            salidaPython = salidaReal
            errMsg = "No se pudo generar el archivo de log del importador."
        Else
            Call BorrarArchivoSiExiste(probePath)
            Call BorrarArchivoSiExiste(logPath)

            cmdExe = Chr(34) & BAT_IMPORTADOR & Chr(34) & " " & _
                     Chr(34) & archivoSubido & Chr(34) & " " & _
                     Chr(34) & logPath & Chr(34)

            cmdFinal = "cmd.exe /c " & Chr(34) & cmdExe & Chr(34)

            salidaCmd = EjecutarComandoRun(cmdFinal)

            If EsperarArchivo(logPath, 15) Then
                salidaReal = LeerArchivoTexto(logPath)
            Else
                salidaReal = "No se pudo leer el archivo de log generado por el BAT." & vbCrLf & _
                             "Ruta esperada: " & logPath & vbCrLf & vbCrLf & _
                             "Salida CMD:" & vbCrLf & salidaCmd
            End If

            salidaPython = "CMD: " & cmdFinal & vbCrLf & _
                           "LOG: " & logPath & vbCrLf & vbCrLf & salidaReal

            facturaID = ExtraerFacturaIDDesdeSalida(salidaReal)

            If facturaID <> "" Then
                msg = "Archivo procesado correctamente. FacturaID = " & facturaID
            ElseIf Left(UCase(Trim(salidaReal)), 3) = "OK|" Then
                msg = "Archivo procesado correctamente."
            Else
                errMsg = "El archivo se subió, pero no se pudo procesar correctamente. Revisá la salida completa más abajo."
            End If
        End If

        Session("CAMION_FileUploaded") = ""
        Session("CAMION_FileUploaded_Name") = ""
    End If
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Camiones - Carga de facturas</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
.container{max-width:1100px;margin:20px auto}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;margin-bottom:20px;background:#fff}
.alert{padding:12px 14px;margin-bottom:15px;border-radius:8px}
.alert-ok{background:#d4edda;color:#155724}
.alert-error{background:#f8d7da;color:#721c24}
.btn{background:#2b7cff;color:#fff;border:none;padding:10px 16px;border-radius:8px;cursor:pointer}
.btn:hover{background:#1f68d8}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;display:inline-block;padding:10px 16px;border-radius:8px}
.dropzone{
    border:2px dashed #9aa7b3;
    border-radius:12px;
    padding:38px 20px;
    text-align:center;
    background:#fafbfc;
    cursor:pointer;
}
.input-file{display:block;margin:14px auto 0 auto}
.nombre-archivo{margin-top:12px;font-weight:bold}
.acciones{margin-top:18px;display:flex;gap:10px;flex-wrap:wrap}
pre{
    margin:0;
    white-space:pre-wrap;
    word-break:break-word;
    font-size:13px;
    line-height:1.45;
    background:#111;
    color:#ddd;
    padding:14px;
    border-radius:10px;
    overflow:auto;
}
.links a{color:#2b7cff;text-decoration:none;font-weight:bold}
</style>

<script>
function toggleSidebar(){
    var sb = document.querySelector('.sidebar');
    if(sb){ sb.classList.toggle('open'); }
}
</script>
</head>

<body>

<!--#include file="header.asp" -->
  
<div class="main-content">

    <div class="d-flex justify-content-between align-items-center mb-4" style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
      <h2 style="margin:0;">Carga de facturas</h2>
      <a href="camiones_facturas.asp" class="btn-sec">Ver facturas</a>
    </div>

    <% If msg <> "" Then %>
        <div class="alert alert-ok"><%=Server.HTMLEncode(msg)%></div>
    <% End If %>

    <% If errMsg <> "" Then %>
        <div class="alert alert-error"><%=Server.HTMLEncode(errMsg)%></div>
    <% End If %>

    <div class="card">
        <h3>Subir PDF de factura</h3>
        <p>Seleccioná el archivo PDF para procesarlo con el importador de camiones.</p>
        <%=htmlUpload%>
    </div>

    <% If nombreArchivoMostrado <> "" Then %>
    <div class="card">
        <h3>Archivo subido</h3>
        <div><strong><%=Server.HTMLEncode(nombreArchivoMostrado)%></strong></div>
    </div>
    <% End If %>

    <% If facturaID <> "" Then %>
    <div class="card links">
        <h3>Resultado</h3>
        <div>Factura generada: <strong><%=facturaID%></strong></div>
        <div style="margin-top:8px;">
            <a href="camiones_factura_control.asp?FacturaID=<%=Server.URLEncode(facturaID)%>">Abrir factura</a>
        </div>
    </div>
    <% End If %>

    <% If salidaPython <> "" Then %>
    <div class="card">
        <h3>Salida del importador</h3>
        <pre><%=Server.HTMLEncode(salidaPython)%></pre>
    </div>
    <% End If %>

</div>

<script>
(function(){
    var input = document.getElementById('File1');
    var nombre = document.getElementById('nombreArchivo');

    if(input){
        input.addEventListener('change', function(){
            if (this.files && this.files.length > 0) {
                nombre.textContent = this.files[0].name;
            } else {
                nombre.textContent = '';
            }
        });
    }
})();
</script>

</body>
</html>
<%
conn.Close
Set conn = Nothing
%>