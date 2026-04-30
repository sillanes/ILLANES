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

Dim CARPETA_UPLOADS
CARPETA_UPLOADS = "C:\Logistica\FacturasCSV"

Dim msg, errMsg, salidaProceso, camionID, archivoSubido, nombreArchivoMostrado
Dim doWhat, step, htmlUpload
msg                   = ""
errMsg                = ""
salidaProceso         = ""
camionID              = ""
archivoSubido         = ""
nombreArchivoMostrado = ""
htmlUpload            = ""

doWhat = CLng("0" & Request("doWhat"))
step   = Trim(Request("step"))

If doWhat = 0 Then
    Session("CAMIONCSV_FileUploaded") = ""
    Session("CAMIONCSV_FileUploaded_Name") = ""
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

Function SafeSql(v)
    SafeSql = Replace("" & Nz(v, ""), "'", "''")
End Function

Function CleanCell(v)
    Dim s
    s = "" & Nz(v, "")
    s = Replace(s, Chr(9), " ")
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim(s)
    CleanCell = s
End Function

Function NormalizeLabel(v)
    Dim s
    s = UCase(CleanCell(v))
    s = Replace(s, "Á", "A")
    s = Replace(s, "É", "E")
    s = Replace(s, "Í", "I")
    s = Replace(s, "Ó", "O")
    s = Replace(s, "Ú", "U")
    s = Replace(s, "Ü", "U")
    s = Replace(s, "Ñ", "N")
    s = Replace(s, ".", "")
    s = Replace(s, ":", "")
    NormalizeLabel = s
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
            If LCase(fso.GetExtensionName(fileObj.Name)) = "csv" Then
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

Function ReadTextFileWithCharset(path, charsetName)
    Dim stm, txt
    txt = ""

    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Mode = 3
    stm.Charset = charsetName
    stm.Open
    stm.LoadFromFile path
    txt = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing

    ReadTextFileWithCharset = txt
End Function

Function ReadTextFileAuto(path)
    On Error Resume Next

    Dim txt
    txt = ""

    ' 1) CSVs de Windows / Excel suelen venir así
    txt = ReadTextFileWithCharset(path, "windows-1252")
    If Err.Number = 0 Then
        ReadTextFileAuto = txt
        Exit Function
    End If
    Err.Clear

    ' 2) Segundo intento UTF-8
    txt = ReadTextFileWithCharset(path, "utf-8")
    If Err.Number = 0 Then
        ReadTextFileAuto = txt
        Exit Function
    End If
    Err.Clear

    ' 3) Último intento Latin1
    txt = ReadTextFileWithCharset(path, "iso-8859-1")
    If Err.Number = 0 Then
        ReadTextFileAuto = txt
        Exit Function
    End If
    Err.Clear

    ReadTextFileAuto = ""
    On Error GoTo 0
End Function

Function GetLineDelimiter(line)
    Dim cComma, cSemi, cTab, i, ch
    cComma = 0
    cSemi = 0
    cTab = 0

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If ch = "," Then cComma = cComma + 1
        If ch = ";" Then cSemi = cSemi + 1
        If ch = vbTab Then cTab = cTab + 1
    Next

    If cTab >= cSemi And cTab >= cComma Then
        GetLineDelimiter = vbTab
    ElseIf cSemi >= cComma Then
        GetLineDelimiter = ";"
    Else
        GetLineDelimiter = ","
    End If
End Function

Function ParseDelimitedLine(line)
    Dim delim, i, ch, inQuotes, token, arr(), idx
    delim = GetLineDelimiter(line)

    ReDim arr(0)
    idx = 0
    token = ""
    inQuotes = False

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(line) And Mid(line, i + 1, 1) = """" Then
                token = token & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = delim And Not inQuotes Then
            arr(idx) = CleanCell(token)
            idx = idx + 1
            ReDim Preserve arr(idx)
            token = ""
        Else
            token = token & ch
        End If
    Next

    arr(idx) = CleanCell(token)
    ParseDelimitedLine = arr
End Function

Function ArrayValue(arr, idx)
    On Error Resume Next
    If IsArray(arr) Then
        If idx >= LBound(arr) And idx <= UBound(arr) Then
            ArrayValue = CleanCell(arr(idx))
        Else
            ArrayValue = ""
        End If
    Else
        ArrayValue = ""
    End If
    On Error GoTo 0
End Function

Function LastNonEmptyValue(arr)
    Dim i
    LastNonEmptyValue = ""
    For i = UBound(arr) To LBound(arr) Step -1
        If Trim(ArrayValue(arr, i)) <> "" Then
            LastNonEmptyValue = ArrayValue(arr, i)
            Exit Function
        End If
    Next
End Function

Function ToSqlNullableText(v)
    If Trim("" & v) = "" Then
        ToSqlNullableText = "NULL"
    Else
        ToSqlNullableText = "'" & SafeSql(v) & "'"
    End If
End Function

Function ToSqlNullableNumber(v)
    Dim s
    s = Trim("" & v)

    If s = "" Then
        ToSqlNullableNumber = "NULL"
        Exit Function
    End If

    s = Replace(s, ".", "")
    s = Replace(s, ",", ".")

    If IsNumeric(s) Then
        ToSqlNullableNumber = Replace(CStr(CDbl(s)), ",", ".")
    Else
        ToSqlNullableNumber = "NULL"
    End If
End Function

Function LooksLikeItemRow(arr)
    Dim c0, c1, c2, c3, c4, c5
    c0 = ArrayValue(arr, 0)
    c1 = ArrayValue(arr, 1)
    c2 = ArrayValue(arr, 2)
    c3 = ArrayValue(arr, 3)
    c4 = ArrayValue(arr, 4)
    c5 = ArrayValue(arr, 5)

    If NormalizeLabel(c0) = "CODIGO EAN" Or NormalizeLabel(c1) = "PRODUCTO" Then
        LooksLikeItemRow = False
        Exit Function
    End If

    If InStr(1, NormalizeLabel(c0), "TOTAL DE BULTOS PARA ESTA EMPRESA", vbTextCompare) > 0 _
       Or InStr(1, NormalizeLabel(c1), "TOTAL DE BULTOS PARA ESTA EMPRESA", vbTextCompare) > 0 Then
        LooksLikeItemRow = False
        Exit Function
    End If

    If c0 <> "" And c1 <> "" And c2 <> "" And c3 <> "" And c4 <> "" Then
        If UCase(c3) = "BULTOS" Or IsNumeric(Replace(c4, ",", ".")) Then
            LooksLikeItemRow = True
            Exit Function
        End If
    End If

    If c0 = "" And c1 <> "" And c2 <> "" And c3 <> "" And c4 <> "" Then
        If UCase(c3) = "BULTOS" Or IsNumeric(Replace(c4, ",", ".")) Then
            LooksLikeItemRow = True
            Exit Function
        End If
    End If

    LooksLikeItemRow = False
End Function

Function ExecScalar(sqlText, fieldName)
    Dim rsTmp, v
    v = ""
    Set rsTmp = conn.Execute(sqlText)
    If Not rsTmp.EOF Then
        v = rsTmp(fieldName)
    End If
    rsTmp.Close
    Set rsTmp = Nothing
    ExecScalar = v
End Function

Sub ExecuteNonQuery(sqlText)
    conn.Execute sqlText
End Sub

Function FindValueAfterLabel(arr, labelText)
    Dim i, lbl
    lbl = NormalizeLabel(labelText)
    FindValueAfterLabel = ""

    For i = LBound(arr) To UBound(arr)
        If NormalizeLabel(ArrayValue(arr, i)) = lbl Then
            If i + 1 <= UBound(arr) Then
                FindValueAfterLabel = ArrayValue(arr, i + 1)
                Exit Function
            End If
        End If
    Next
End Function

Function do_Upload_CSV_Camiones()
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

        Form.Files.Save CARPETA_UPLOADS

        archivoDetectado = ObtenerArchivoMasReciente(CARPETA_UPLOADS)

        If Trim(archivoDetectado) = "" Then
            errMsg = "No se encontró el archivo CSV luego de la subida."
            hResult = "<div class='alert alert-error'>" & errMsg & "</div>"
        Else
            If LCase(ObtenerExtension(archivoDetectado)) <> "csv" Then
                errMsg = "Solo se permiten archivos CSV."
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

                Session("CAMIONCSV_FileUploaded") = archivoFinal
                Session("CAMIONCSV_FileUploaded_Name") = Mid(archivoFinal, InStrRev(archivoFinal, "\") + 1)

                archivoSubido = Session("CAMIONCSV_FileUploaded")
                nombreArchivoMostrado = Session("CAMIONCSV_FileUploaded_Name")

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
    HTML = HTML & "  <div><strong>Arrastrá acá el CSV</strong> o hacé click para seleccionarlo</div>"
    HTML = HTML & "  <div style='margin-top:8px;color:#666;'>Solo CSV de carga de camiones</div>"
    HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required accept='.csv,text/csv'>"
    HTML = HTML & "  <div class='nombre-archivo' id='nombreArchivo'></div>"
    HTML = HTML & "</div>"

    HTML = HTML & "<div class='acciones'>"
    HTML = HTML & "  <button type='submit' class='btn'>Subir CSV</button>"
    HTML = HTML & "</div>"
    HTML = HTML & "</form>"

    HTML = HTML & "<script>" & vbCrLf
    HTML = HTML & "function isMobile(){ return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent); }" & vbCrLf
    HTML = HTML & "function ProgressBarSmart(){ if(isMobile()){ return true; }" & vbCrLf
    HTML = HTML & "  var ProgressURL='progress.asp?UploadID=" & UploadID & "';" & vbCrLf
    HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
    HTML = HTML & "  return true; }" & vbCrLf
    HTML = HTML & "</script>"

    do_Upload_CSV_Camiones = HTML
End Function

Sub ProcesarCSV(pathArchivo)
    On Error Resume Next

    Dim rawText, normalizedText, lines, lineCount
    Dim i, lineText, arr, sql
    Dim headerLine(17)
    Dim nroCarga, fechaCarga, nroCuenta, empresaTransporte, chofer, telefono, cantidadPallets
    Dim ordenCompra, empresa, nroFactura, nroRemito, totalBultos, cabeceraID
    Dim codigoEAN, producto, descripcion, unidadMedida, cantidad
    Dim filasCabecera, filasItems

    rawText = ReadTextFileAuto(pathArchivo)
    normalizedText = Replace(rawText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    lines = Split(normalizedText, vbLf)
    lineCount = UBound(lines)

    For i = 1 To 17
        headerLine(i) = ""
    Next

    nroCarga = ""
    fechaCarga = ""
    nroCuenta = ""
    empresaTransporte = ""
    chofer = ""
    telefono = ""
    cantidadPallets = ""

    For i = 0 To lineCount
        If i <= 16 Then
            headerLine(i + 1) = Trim("" & lines(i))
        End If

        lineText = Trim("" & lines(i))
        If lineText <> "" Then
            arr = ParseDelimitedLine(lineText)

            Select Case NormalizeLabel(ArrayValue(arr, 0))
                Case "NRO DE CARGA"
                    nroCarga = ArrayValue(arr, 1)
                Case "FECHA"
                    fechaCarga = ArrayValue(arr, 1)
                Case "NRO DE CUENTA"
                    nroCuenta = ArrayValue(arr, 1)
                Case "EMPRESA DE TRANSPORTE"
                    empresaTransporte = ArrayValue(arr, 1)
                Case "CHOFER"
                    chofer = ArrayValue(arr, 1)
                Case "TELEFONO"
                    telefono = ArrayValue(arr, 1)
                Case "CANTIDAD DE PALLETS"
                    cantidadPallets = ArrayValue(arr, 1)
            End Select
        End If
    Next

    sql = "EXEC dbo.usp_Camiones_Ins " & _
          "@UsuarioImportacion=" & ToSqlNullableText(Session("currentUser")) & ", " & _
          "@ArchivoNombre=" & ToSqlNullableText(nombreArchivoMostrado) & ", " & _
          "@NroCarga=" & ToSqlNullableText(nroCarga) & ", " & _
          "@FechaCarga=" & ToSqlNullableText(fechaCarga) & ", " & _
          "@NroCuenta=" & ToSqlNullableText(nroCuenta) & ", " & _
          "@EmpresaTransporte=" & ToSqlNullableText(empresaTransporte) & ", " & _
          "@Chofer=" & ToSqlNullableText(chofer) & ", " & _
          "@Telefono=" & ToSqlNullableText(telefono) & ", " & _
          "@CantidadPallets=" & ToSqlNullableText(cantidadPallets) & ", " & _
          "@CabeceraLinea01=" & ToSqlNullableText(headerLine(1)) & ", " & _
          "@CabeceraLinea02=" & ToSqlNullableText(headerLine(2)) & ", " & _
          "@CabeceraLinea03=" & ToSqlNullableText(headerLine(3)) & ", " & _
          "@CabeceraLinea04=" & ToSqlNullableText(headerLine(4)) & ", " & _
          "@CabeceraLinea05=" & ToSqlNullableText(headerLine(5)) & ", " & _
          "@CabeceraLinea06=" & ToSqlNullableText(headerLine(6)) & ", " & _
          "@CabeceraLinea07=" & ToSqlNullableText(headerLine(7)) & ", " & _
          "@CabeceraLinea08=" & ToSqlNullableText(headerLine(8)) & ", " & _
          "@CabeceraLinea09=" & ToSqlNullableText(headerLine(9)) & ", " & _
          "@CabeceraLinea10=" & ToSqlNullableText(headerLine(10)) & ", " & _
          "@CabeceraLinea11=" & ToSqlNullableText(headerLine(11)) & ", " & _
          "@CabeceraLinea12=" & ToSqlNullableText(headerLine(12)) & ", " & _
          "@CabeceraLinea13=" & ToSqlNullableText(headerLine(13)) & ", " & _
          "@CabeceraLinea14=" & ToSqlNullableText(headerLine(14)) & ", " & _
          "@CabeceraLinea15=" & ToSqlNullableText(headerLine(15)) & ", " & _
          "@CabeceraLinea16=" & ToSqlNullableText(headerLine(16)) & ", " & _
          "@CabeceraLinea17=" & ToSqlNullableText(headerLine(17))

    camionID = ExecScalar(sql, "CamionID")

    If Err.Number <> 0 Then
        errMsg = "Error al guardar la cabecera general del CSV: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    ordenCompra = ""
    empresa = ""
    nroFactura = ""
    nroRemito = ""
    totalBultos = ""
    cabeceraID = 0
    filasCabecera = 0
    filasItems = 0

    For i = 0 To lineCount
        lineText = Trim("" & lines(i))
        If lineText <> "" Then
            arr = ParseDelimitedLine(lineText)

            If NormalizeLabel(ArrayValue(arr, 0)) = "ORDEN DE COMPRA" Then
                ordenCompra = ArrayValue(arr, 1)
                empresa = ""
                nroFactura = ""
                nroRemito = ""
                totalBultos = ""
                cabeceraID = 0

            ElseIf NormalizeLabel(ArrayValue(arr, 0)) = "EMPRESA" Then
                empresa = ArrayValue(arr, 1)

            ElseIf FindValueAfterLabel(arr, "Nro. de Factura") <> "" _
                Or FindValueAfterLabel(arr, "Nro de Factura") <> "" _
                Or FindValueAfterLabel(arr, "Nro. de Remito") <> "" _
                Or FindValueAfterLabel(arr, "Nro de Remito") <> "" Then

                nroFactura = FindValueAfterLabel(arr, "Nro. de Factura")
                If nroFactura = "" Then nroFactura = FindValueAfterLabel(arr, "Nro de Factura")

                nroRemito = FindValueAfterLabel(arr, "Nro. de Remito")
                If nroRemito = "" Then nroRemito = FindValueAfterLabel(arr, "Nro de Remito")

            ElseIf InStr(1, NormalizeLabel(ArrayValue(arr, 1)), "TOTAL DE BULTOS PARA ESTA EMPRESA", vbTextCompare) > 0 Or _
                   InStr(1, NormalizeLabel(ArrayValue(arr, 0)), "TOTAL DE BULTOS PARA ESTA EMPRESA", vbTextCompare) > 0 Then

                totalBultos = LastNonEmptyValue(arr)

                If cabeceraID <> "" And CLng("0" & cabeceraID) > 0 Then
                    sql = "UPDATE dbo.Camiones_Cabecera " & _
                          "SET TotalBultosEmpresa = " & ToSqlNullableNumber(totalBultos) & " " & _
                          "WHERE CabeceraID = " & CLng(cabeceraID)
                    ExecuteNonQuery sql
                End If

            ElseIf LooksLikeItemRow(arr) Then
                If CLng("0" & cabeceraID) = 0 Then
                    sql = "EXEC dbo.usp_Camiones_Cabecera_Ins " & _
                          "@CamionID=" & CLng("0" & camionID) & ", " & _
                          "@OrdenCompra=" & ToSqlNullableText(ordenCompra) & ", " & _
                          "@Empresa=" & ToSqlNullableText(empresa) & ", " & _
                          "@NroFactura=" & ToSqlNullableText(nroFactura) & ", " & _
                          "@NroRemito=" & ToSqlNullableText(nroRemito) & ", " & _
                          "@TotalBultosEmpresa=NULL"

                    cabeceraID = ExecScalar(sql, "CabeceraID")
                    filasCabecera = filasCabecera + 1
                End If

                If ArrayValue(arr, 0) = "" And ArrayValue(arr, 1) <> "" Then
                    codigoEAN    = ""
                    producto     = ArrayValue(arr, 1)
                    descripcion  = ArrayValue(arr, 2)
                    unidadMedida = ArrayValue(arr, 3)
                    cantidad     = ArrayValue(arr, 4)
                Else
                    codigoEAN    = ArrayValue(arr, 0)
                    producto     = ArrayValue(arr, 1)
                    descripcion  = ArrayValue(arr, 2)
                    unidadMedida = ArrayValue(arr, 3)
                    cantidad     = ArrayValue(arr, 4)
                End If

                sql = "EXEC dbo.usp_Camiones_Cabecera_Items_Ins " & _
                      "@CabeceraID=" & CLng("0" & cabeceraID) & ", " & _
                      "@CodigoEAN=" & ToSqlNullableText(codigoEAN) & ", " & _
                      "@Producto=" & ToSqlNullableText(producto) & ", " & _
                      "@Descripcion=" & ToSqlNullableText(descripcion) & ", " & _
                      "@UnidadMedida=" & ToSqlNullableText(unidadMedida) & ", " & _
                      "@Cantidad=" & ToSqlNullableNumber(cantidad)

                ExecuteNonQuery sql
                filasItems = filasItems + 1
            End If
        End If
    Next

    If Err.Number <> 0 Then
        errMsg = "Error al procesar el detalle del CSV: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    salidaProceso = "Archivo: " & nombreArchivoMostrado & vbCrLf & _
                    "CamionID: " & camionID & vbCrLf & _
                    "Facturas importadas: " & filasCabecera & vbCrLf & _
                    "Items importados: " & filasItems

    msg = "CSV procesado correctamente. CamionID = " & camionID

    On Error GoTo 0
End Sub

Call CrearCarpetaSiNoExiste(CARPETA_UPLOADS)
htmlUpload = do_Upload_CSV_Camiones()

If Trim(Session("CAMIONCSV_FileUploaded") & "") <> "" Then
    archivoSubido = Session("CAMIONCSV_FileUploaded")
    nombreArchivoMostrado = Session("CAMIONCSV_FileUploaded_Name")

    If archivoSubido <> "" Then
        Call ProcesarCSV(archivoSubido)

        Session("CAMIONCSV_FileUploaded") = ""
        Session("CAMIONCSV_FileUploaded_Name") = ""
    End If
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Camiones - Carga de CSV</title>

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
.btn{background:#198754;color:#fff;border:none;padding:10px 16px;border-radius:8px;cursor:pointer}
.btn:hover{background:#157347}
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
      <h2 style="margin:0;">Carga de CSV</h2>
      <div style="display:flex;gap:10px;flex-wrap:wrap;">
        <a href="camiones_facturas.asp" class="btn-sec">Ver facturas</a>
        <a href="camiones_facturas_carga.asp" class="btn-sec">Cargar PDF</a>
      </div>
    </div>

    <% If msg <> "" Then %>
        <div class="alert alert-ok"><%=Server.HTMLEncode(msg)%></div>
    <% End If %>

    <% If errMsg <> "" Then %>
        <div class="alert alert-error"><%=Server.HTMLEncode(errMsg)%></div>
    <% End If %>

    <div class="card">
        <h3>Subir CSV</h3>
        <p>Seleccioná el archivo CSV para importar la cabecera general, las facturas y sus ítems.</p>
        <%=htmlUpload%>
    </div>

    <% If nombreArchivoMostrado <> "" Then %>
    <div class="card">
        <h3>Archivo subido</h3>
        <div><strong><%=Server.HTMLEncode(nombreArchivoMostrado)%></strong></div>
    </div>
    <% End If %>

    <% If camionID <> "" Then %>
    <div class="card">
        <h3>Resultado</h3>
        <div>CamionID generado: <strong><%=Server.HTMLEncode(camionID)%></strong></div>
    </div>
    <% End If %>

    <% If salidaProceso <> "" Then %>
    <div class="card">
        <h3>Salida del proceso</h3>
        <pre><%=Server.HTMLEncode(salidaProceso)%></pre>
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