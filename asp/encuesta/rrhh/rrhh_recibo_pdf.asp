<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True
Server.ScriptTimeout = 600

If Trim(Session("currentUser") & "") <> "admin" AND Trim(Session("currentUser") & "") <> "rrhh"  Then
    Response.Redirect "../login.asp"
End If

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function ToInt(v, alt)
    On Error Resume Next
    Dim n
    n = alt
    If Trim(CStr(v & "")) <> "" Then
        n = CLng(v)
        If Err.Number <> 0 Then
            n = alt
            Err.Clear
        End If
    End If
    ToInt = n
    On Error GoTo 0
End Function

Function EsUrlAbsoluta(s)
    s = LCase(Trim(CStr(Nz(s, ""))))
    EsUrlAbsoluta = (Left(s, 7) = "http://" Or Left(s, 8) = "https://")
End Function

Function EsRutaWeb(s)
    s = Trim(CStr(Nz(s, "")))
    EsRutaWeb = False
    If s <> "" Then
        If Left(s, 1) = "/" Then EsRutaWeb = True
    End If
End Function

Sub MostrarError(msg)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - PDF</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
.box{
    max-width:700px;
    margin:40px auto;
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:12px;
    padding:20px;
    font-family:Arial, Helvetica, sans-serif;
}
.err{
    background:#fef2f2;
    border:1px solid #fecaca;
    color:#991b1b;
    padding:12px 14px;
    border-radius:10px;
}
</style>
</head>
<body>
<div class="box">
    <h2>PDF no disponible</h2>
    <div class="err"><%=Server.HTMLEncode(msg)%></div>
</div>
</body>
</html>
<%
End Sub

Dim reciboID, tipo
reciboID = ToInt(Request.QueryString("reciboid"), 0)
tipo     = LCase(Trim(CStr(Request.QueryString("tipo") & "")))

If reciboID <= 0 Then
    Call MostrarError("Recibo inválido.")
    Response.End
End If

If tipo <> "original" And tipo <> "firmado" Then
    tipo = "original"
End If

Dim rsCheck, sqlTieneFirmado, tieneCampoFirmado
tieneCampoFirmado = False

sqlTieneFirmado = "SELECT COUNT(*) AS Cnt " & _
                  "FROM INFORMATION_SCHEMA.COLUMNS " & _
                  "WHERE TABLE_SCHEMA='rrhh' AND TABLE_NAME='Recibos' AND COLUMN_NAME='RutaArchivoFirmado'"

Set rsCheck = conn.Execute(sqlTieneFirmado)
If Not rsCheck.EOF Then
    If CLng(Nz(rsCheck("Cnt"),0)) > 0 Then tieneCampoFirmado = True
End If
rsCheck.Close
Set rsCheck = Nothing

Dim sql, rs, rutaOriginal, rutaFirmado, rutaFinal
If tieneCampoFirmado Then
    sql = "SELECT RutaArchivo, RutaArchivoFirmado FROM rrhh.Recibos WHERE ReciboID=" & reciboID
Else
    sql = "SELECT RutaArchivo, CAST(NULL AS VARCHAR(500)) AS RutaArchivoFirmado FROM rrhh.Recibos WHERE ReciboID=" & reciboID
End If

Set rs = conn.Execute(sql)

If rs.EOF Then
    rs.Close : Set rs = Nothing
    Call MostrarError("No se encontró el recibo indicado.")
    Response.End
End If

rutaOriginal = Trim(CStr(Nz(rs("RutaArchivo"), "")))
rutaFirmado  = Trim(CStr(Nz(rs("RutaArchivoFirmado"), "")))

rs.Close
Set rs = Nothing

If tipo = "firmado" Then
    rutaFinal = rutaFirmado
Else
    rutaFinal = rutaOriginal
End If

If Trim(rutaFinal) = "" Then
    If tipo = "firmado" Then
        Call MostrarError("Este recibo todavía no tiene PDF firmado.")
    Else
        Call MostrarError("Este recibo no tiene ruta de PDF original.")
    End If
    Response.End
End If

If EsUrlAbsoluta(rutaFinal) Or EsRutaWeb(rutaFinal) Then
    Response.Redirect rutaFinal
End If

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(rutaFinal) Then
    Set fso = Nothing
    Call MostrarError("El archivo no existe en la ruta configurada: " & rutaFinal)
    Response.End
End If

Dim stm
Set stm = Server.CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.LoadFromFile rutaFinal

Response.Clear
Response.ContentType = "application/pdf"
Response.AddHeader "Content-Disposition", "inline; filename=recibo_" & reciboID & "_" & tipo & ".pdf"
Response.BinaryWrite stm.Read

stm.Close
Set stm = Nothing
Set fso = Nothing

conn.Close
Set conn = Nothing
Response.End
%>