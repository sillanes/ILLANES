<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 3600

If CLng("0" & Session("Empleado_UsuarioID")) <= 0 Or CLng("0" & Session("Empleado_EmpleadoID")) <= 0 Then
    Response.Redirect "empleado_login.asp"
End If

Dim CARPETA_LOGS, CARPETA_TEST, BAT_TEST
CARPETA_LOGS = "C:\RRHH\Logs\"
CARPETA_TEST = "C:\RRHH\Archivos\RecibosFirmados\Test\"
BAT_TEST     = "C:\RRHH\Firmador\run_firmar_recibo_test.bat"

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

Function EjecutarComandoRun(cmd)
    On Error Resume Next
    Dim sh, rc

    Set sh = Server.CreateObject("WScript.Shell")
    rc = sh.Run(cmd, 0, True)

    If Err.Number <> 0 Then
        EjecutarComandoRun = "ERROR al ejecutar comando: " & Err.Description & vbCrLf & "CMD: " & cmd
        Err.Clear
    Else
        EjecutarComandoRun = "ExitCode=" & rc
    End If

    Set sh = Nothing
    On Error GoTo 0
End Function

Function GenerarNombreRandom(prefijo, extension)
    Randomize
    GenerarNombreRandom = prefijo & "_" & _
        Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & "_" & _
        Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2) & "_" & _
        CStr(Int((9999 - 1000 + 1) * Rnd + 1000)) & extension
End Function

Function ExtraerValorLog(texto, clave)
    Dim re, m
    ExtraerValorLog = ""

    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = False
    re.Pattern = clave & "\s*=\s*(.+)"

    If re.Test(texto) Then
        Set m = re.Execute(texto)
        ExtraerValorLog = Trim(m(0).SubMatches(0))
    End If

    Set re = Nothing
End Function

Dim empleadoID, reciboid, accion
Dim nombreFirmante, fechaFirma
Dim pageIndex, firmaX, firmaY, firmaSize
Dim aclarX, aclarY, aclarSize
Dim fechaX, fechaY, fechaSize
Dim msgOk, msgErr, salidaLog

empleadoID = CLng("0" & Session("Empleado_EmpleadoID"))
reciboid   = CLng("0" & Request("reciboid"))
accion     = LCase(Trim(Request("accion")))

nombreFirmante = Trim(Request("nombrefirmante"))
fechaFirma     = Trim(Request("fechafirma"))

pageIndex = Trim(Request("pageindex"))
If pageIndex = "" Then pageIndex = "0"

firmaX = Trim(Request("firmax"))
If firmaX = "" Then firmaX = "330"

firmaY = Trim(Request("firmay"))
If firmaY = "" Then firmaY = "115"

firmaSize = Trim(Request("firmasize"))
If firmaSize = "" Then firmaSize = "10"

aclarX = Trim(Request("aclarx"))
If aclarX = "" Then aclarX = "300"

aclarY = Trim(Request("aclary"))
If aclarY = "" Then aclarY = "100"

aclarSize = Trim(Request("aclarsize"))
If aclarSize = "" Then aclarSize = "9"

fechaX = Trim(Request("fechax"))
If fechaX = "" Then fechaX = "300"

fechaY = Trim(Request("fechay"))
If fechaY = "" Then fechaY = "85"

fechaSize = Trim(Request("fechasize"))
If fechaSize = "" Then fechaSize = "8"

If nombreFirmante = "" Then nombreFirmante = Trim("" & Session("Empleado_Nombre"))
If fechaFirma = "" Then
    fechaFirma = Year(Now()) & "-" & Right("0" & Month(Now()),2) & "-" & Right("0" & Day(Now()),2)
End If

msgOk = ""
msgErr = ""
salidaLog = ""

If reciboid <= 0 Then
    Response.Write "ReciboID inválido."
    Response.End
End If

Dim rs
Set rs = conn.Execute("EXEC empleado.usp_Empleado_Recibo_PDF_Sel " & empleadoID & ", " & reciboid)

If rs.EOF Then
    Response.Write "No existe el recibo o no pertenece al empleado."
    Response.End
End If

If accion = "probar" Then
    Dim rutaEntrada, rutaSalida, rutaLog, salidaCmd, pdfTest, cmdExe, cmdFinal

    rutaEntrada = Trim("" & rs("RutaArchivo"))

    If rutaEntrada = "" Then
        msgErr = "El recibo no tiene PDF asociado."
        Session("Empleado_TestPDFPath") = ""
    Else
        Call CrearCarpetaSiNoExiste(CARPETA_LOGS)
        Call CrearCarpetaSiNoExiste(CARPETA_TEST)

        rutaSalida = CARPETA_TEST & GenerarNombreRandom("test_firma_" & reciboid, ".pdf")
        rutaLog    = CARPETA_LOGS & GenerarNombreRandom("test_firma_" & reciboid, ".txt")

        cmdExe = Chr(34) & BAT_TEST & Chr(34) & " " & _
                 Chr(34) & rutaEntrada & Chr(34) & " " & _
                 Chr(34) & rutaSalida & Chr(34) & " " & _
                 Chr(34) & nombreFirmante & Chr(34) & " " & _
                 Chr(34) & fechaFirma & Chr(34) & " " & _
                 Chr(34) & pageIndex & Chr(34) & " " & _
                 Chr(34) & firmaX & Chr(34) & " " & _
                 Chr(34) & firmaY & Chr(34) & " " & _
                 Chr(34) & firmaSize & Chr(34) & " " & _
                 Chr(34) & aclarX & Chr(34) & " " & _
                 Chr(34) & aclarY & Chr(34) & " " & _
                 Chr(34) & aclarSize & Chr(34) & " " & _
                 Chr(34) & fechaX & Chr(34) & " " & _
                 Chr(34) & fechaY & Chr(34) & " " & _
                 Chr(34) & fechaSize & Chr(34) & " " & _
                 Chr(34) & rutaLog & Chr(34)

        cmdFinal = "cmd.exe /c " & Chr(34) & cmdExe & Chr(34)
        salidaCmd = EjecutarComandoRun(cmdFinal)

        If EsperarArchivo(rutaLog, 15) Then
            salidaLog = LeerArchivoTexto(rutaLog)
        Else
            salidaLog = "No se pudo leer el log de prueba." & vbCrLf & _
                        "CMD=" & cmdFinal & vbCrLf & _
                        salidaCmd
        End If

        pdfTest = ExtraerValorLog(salidaLog, "PDF_TEST")

        If pdfTest <> "" Then
            Session("Empleado_TestPDFPath") = pdfTest
            msgOk = "PDF de prueba generado."
            salidaLog = "CMD=" & cmdFinal & vbCrLf & vbCrLf & salidaLog
        Else
            Session("Empleado_TestPDFPath") = ""
            msgErr = "No se pudo generar el PDF de prueba."
            salidaLog = "CMD=" & cmdFinal & vbCrLf & vbCrLf & salidaLog
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Prueba de coordenadas de firma</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<style>
body{margin:0;font-family:Arial, Helvetica, sans-serif;background:#f4f6f9}
.container{max-width:1300px;margin:20px auto;padding:0 12px;box-sizing:border-box}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;background:#fff;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:12px}
.campo label{display:block;font-size:12px;color:#555;margin-bottom:6px}
.campo input{width:100%;padding:10px;border:1px solid #ccc;border-radius:8px;box-sizing:border-box}
.btn{background:#2b7cff;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.alert-ok{background:#d4edda;color:#155724;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.alert-error{background:#f8d7da;color:#721c24;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.pdf-frame{width:100%;height:75vh;border:1px solid #ddd;border-radius:10px;background:#fff}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}
pre{margin:0;white-space:pre-wrap;word-break:break-word;font-size:12px;background:#111;color:#ddd;padding:14px;border-radius:10px;overflow:auto}

@media (max-width: 900px){
    .grid{grid-template-columns:repeat(2,minmax(160px,1fr))}
}
@media (max-width: 768px){
    .container{margin:12px auto;padding:0 10px}
    .card{padding:14px}
    .grid{grid-template-columns:1fr}
    .actions a,.actions button{width:100%;text-align:center;box-sizing:border-box}
    .pdf-frame{height:60vh}
}
</style>
</head>
<body>
<div class="container">

    <div class="card">
        <h2 style="margin-top:0;">Prueba de coordenadas de firma</h2>
        <div>ReciboID: <strong><%=reciboid%></strong></div>
        <div>Archivo original: <strong><%=Server.HTMLEncode(Nz(rs("NombreArchivo"), ""))%></strong></div>
        <div>Ruta original: <strong><%=Server.HTMLEncode(Nz(rs("RutaArchivo"), ""))%></strong></div>
    </div>

    <% If msgOk <> "" Then %>
        <div class="alert-ok"><%=Server.HTMLEncode(msgOk)%></div>
    <% End If %>

    <% If msgErr <> "" Then %>
        <div class="alert-error"><%=Server.HTMLEncode(msgErr)%></div>
    <% End If %>

    <div class="card">
        <form method="post" action="empleado_recibo_firmar_test.asp">
            <input type="hidden" name="reciboid" value="<%=reciboid%>">
            <input type="hidden" name="accion" value="probar">

            <div class="grid">
                <div class="campo">
                    <label>Página (base 0)</label>
                    <input type="text" name="pageindex" value="<%=Server.HTMLEncode(pageIndex)%>">
                </div>
                <div class="campo">
                    <label>Nombre firmante</label>
                    <input type="text" name="nombrefirmante" value="<%=Server.HTMLEncode(nombreFirmante)%>">
                </div>
                <div class="campo">
                    <label>Fecha firma</label>
                    <input type="text" name="fechafirma" value="<%=Server.HTMLEncode(fechaFirma)%>">
                </div>
                <div class="campo">
                    <label>ReciboID</label>
                    <input type="text" value="<%=reciboid%>" disabled>
                </div>

                <div class="campo">
                    <label>Firma X</label>
                    <input type="text" name="firmax" value="<%=Server.HTMLEncode(firmaX)%>">
                </div>
                <div class="campo">
                    <label>Firma Y</label>
                    <input type="text" name="firmay" value="<%=Server.HTMLEncode(firmaY)%>">
                </div>
                <div class="campo">
                    <label>Firma Size</label>
                    <input type="text" name="firmasize" value="<%=Server.HTMLEncode(firmaSize)%>">
                </div>
                <div class="campo">
                    <label>Aclaración X</label>
                    <input type="text" name="aclarx" value="<%=Server.HTMLEncode(aclarX)%>">
                </div>

                <div class="campo">
                    <label>Aclaración Y</label>
                    <input type="text" name="aclary" value="<%=Server.HTMLEncode(aclarY)%>">
                </div>
                <div class="campo">
                    <label>Aclaración Size</label>
                    <input type="text" name="aclarsize" value="<%=Server.HTMLEncode(aclarSize)%>">
                </div>
                <div class="campo">
                    <label>Fecha X</label>
                    <input type="text" name="fechax" value="<%=Server.HTMLEncode(fechaX)%>">
                </div>
                <div class="campo">
                    <label>Fecha Y</label>
                    <input type="text" name="fechay" value="<%=Server.HTMLEncode(fechaY)%>">
                </div>

                <div class="campo">
                    <label>Fecha Size</label>
                    <input type="text" name="fechasize" value="<%=Server.HTMLEncode(fechaSize)%>">
                </div>
            </div>

            <div class="actions">
                <button type="submit" class="btn">Generar PDF de prueba</button>
                <a href="empleado_recibo_detalle.asp?reciboid=<%=reciboid%>" class="btn-sec">Volver</a>
            </div>
        </form>
    </div>

    <% If msgOk <> "" And Trim("" & Session("Empleado_TestPDFPath")) <> "" Then %>
    <div class="card">
        <h3 style="margin-top:0;">Vista previa PDF de prueba</h3>
        <iframe class="pdf-frame" src="empleado_recibo_pdf_test.asp"></iframe>
        <div class="actions">
            <a href="empleado_recibo_pdf_test.asp" target="_blank" class="btn-sec">Abrir prueba en otra pestaña</a>
        </div>
    </div>
    <% End If %>

    <% If salidaLog <> "" Then %>
    <div class="card">
        <h3 style="margin-top:0;">Salida de prueba</h3>
        <pre><%=Server.HTMLEncode(salidaLog)%></pre>
    </div>
    <% End If %>

</div>
</body>
</html>
<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>