<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True
Server.ScriptTimeout = 600

If CLng("0" & Session("Empleado_UsuarioID")) <= 0 Or CLng("0" & Session("Empleado_EmpleadoID")) <= 0 Then
    Response.Status = "403 Forbidden"
    Response.Write "Acceso denegado."
    Response.End
End If

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

Dim empleadoID, reciboid, modo
empleadoID = CLng("0" & Session("Empleado_EmpleadoID"))
reciboid   = CLng("0" & Request.QueryString("reciboid"))
modo       = LCase(Trim(Request.QueryString("modo")))

If reciboid <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "ReciboID inválido."
    Response.End
End If

Dim rs
Dim rutaArchivoDB, nombreArchivo
Dim rutaArchivoFirmadoDB, nombreArchivoFirmado
Dim rutaArchivoOriginal
Dim rutaElegidaDB, nombreElegido

Set rs = conn.Execute("EXEC empleado.usp_Empleado_Recibo_PDF_Sel " & empleadoID & ", " & reciboid)

If rs.EOF Then
    Response.Status = "404 Not Found"
    Response.Write "No existe el recibo."
    Response.End
End If

rutaArchivoDB         = Trim(CStr(Nz(rs("RutaArchivo"), "")))
nombreArchivo         = Trim(CStr(Nz(rs("NombreArchivo"), "")))

rutaArchivoFirmadoDB  = Trim(CStr(Nz(rs("RutaArchivoFirmado"), "")))
nombreArchivoFirmado  = Trim(CStr(Nz(rs("NombreArchivoFirmado"), "")))

rutaArchivoOriginal   = Trim(CStr(Nz(rs("RutaArchivoOriginal"), "")))

rutaElegidaDB = ""
nombreElegido = ""

Select Case modo
    Case "original"
        If rutaArchivoOriginal <> "" Then
            rutaElegidaDB = rutaArchivoOriginal
        Else
            rutaElegidaDB = rutaArchivoDB
        End If
        nombreElegido = nombreArchivo

    Case "firmado"
        rutaElegidaDB = rutaArchivoFirmadoDB
        nombreElegido = nombreArchivoFirmado

    Case Else
        If rutaArchivoFirmadoDB <> "" Then
            rutaElegidaDB = rutaArchivoFirmadoDB
            nombreElegido = nombreArchivoFirmado
        ElseIf rutaArchivoDB <> "" Then
            rutaElegidaDB = rutaArchivoDB
            nombreElegido = nombreArchivo
        ElseIf rutaArchivoOriginal <> "" Then
            rutaElegidaDB = rutaArchivoOriginal
            nombreElegido = nombreArchivo
        End If
End Select

If Trim(rutaElegidaDB) = "" Then
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

    Response.Status = "404 Not Found"

    If modo = "firmado" Then
        Response.Write "No hay PDF firmado asociado."
    Else
        Response.Write "No hay archivo asociado."
    End If

    Response.End
End If

If Trim(nombreElegido) = "" Then
    nombreElegido = "recibo_" & reciboid & ".pdf"
End If

' Marca lectura al abrir cualquier versión del PDF
conn.Execute "EXEC empleado.usp_Empleado_Recibo_MarcarLeido " & empleadoID & ", " & reciboid

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

Dim rutaUNC, rutaLocal, rutaFinal
rutaUNC   = LocalToUNC(rutaElegidaDB)
rutaLocal = UNCToLocal(rutaElegidaDB)
rutaFinal = ""

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If fso.FileExists(rutaLocal) Then
    rutaFinal = rutaLocal
ElseIf fso.FileExists(rutaUNC) Then
    rutaFinal = rutaUNC
End If

If rutaFinal = "" Then
    Set fso = Nothing
    Response.Status = "404 Not Found"
    Response.Write "No se encontró el archivo PDF asociado al recibo."
    Response.End
End If

Set fso = Nothing

Dim stm
Set stm = Server.CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.LoadFromFile rutaFinal

Response.Clear
Response.ContentType = "application/pdf"
Response.AddHeader "Content-Disposition", "inline; filename=""" & nombreElegido & """"
Response.BinaryWrite stm.Read

stm.Close
Set stm = Nothing
Response.End
%>