<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Response.ContentType = "text/plain"
Response.Write "=== DEBUG FORM UPLOAD ===" & vbCrLf



' ====================================================
' FUNCIONES AUXILIARES
' ====================================================

' Guardar binarios en archivo físico
Sub SaveBinaryFile(filePath, fileBytes)
    Dim ado
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1 ' adTypeBinary
    ado.Open
    ado.Write fileBytes
    ado.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    ado.Close
    Set ado = Nothing
End Sub

' Convertir binario a string
Function BinaryToString(Binary)
    Dim ado
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1
    ado.Open
    ado.Write Binary
    ado.Position = 0
    ado.Type = 2
    ado.Charset = "UTF-8"
    BinaryToString = ado.ReadText
    ado.Close
    Set ado = Nothing
End Function

' ====================================================
' PROCESAR UPLOAD
' ====================================================

Dim bytesTotal, binData, strData, boundary
bytesTotal = Request.TotalBytes
binData = Request.BinaryRead(bytesTotal)
strData = BinaryToString(binData)

 
Response.Write "Total bytes: " & bytesTotal & vbCrLf
Response.Write "Primeros 2000 caracteres:" & vbCrLf
Response.Write Left(strData,2000)
Response.End


' El boundary es la primera línea del request
boundary = Left(strData, InStr(strData, vbCrLf) - 1)

' ====================================================
' EXTRAER CAMPOS DE FORMULARIO
' ====================================================
Function GetFormValue(fieldName, strData)
    Dim posField, posValueStart, posValueEnd, value
    GetFormValue = ""
    posField = InStr(1, strData, "name=""" & fieldName & """", vbTextCompare)
    If posField > 0 Then
        posValueStart = InStr(posField, strData, vbCrLf & vbCrLf) + 4
        posValueEnd = InStr(posValueStart, strData, boundary) - 2
        value = Mid(strData, posValueStart, posValueEnd - posValueStart)
        GetFormValue = Trim(value)
    End If
End Function

Dim NombreCampania, Descripcion, Canal, MensajeWhatsApp, TemplateWhatsApp, AsuntoEmail, CuerpoEmail
NombreCampania   = GetFormValue("NombreCampania", strData)
Descripcion      = GetFormValue("Descripcion", strData)
Canal            = GetFormValue("Canal", strData) ' "WhatsApp" o "Email"
MensajeWhatsApp  = GetFormValue("MensajeWhatsApp", strData)
TemplateWhatsApp = GetFormValue("TemplateWhatsApp", strData)
AsuntoEmail      = GetFormValue("MensajeEmailAsunto", strData)
CuerpoEmail      = GetFormValue("MensajeEmailCuerpo", strData)

' ====================================================
' EXTRAER ARCHIVO SUBIDO (Excel)
' ====================================================
Dim posFile, posStart, posEnd, fileData, fileName, savePath

fileName = ""
Set fileData = Nothing

posFile = InStr(1, strData, "filename=""")
If posFile > 0 Then
    ' Extraer nombre
    fileName = Mid(strData, posFile + 10, InStr(posFile + 10, strData, """") - (posFile + 10))
    If fileName <> "" Then
        ' Donde empieza el binario
        posStart = InStr(posFile, strData, vbCrLf & vbCrLf) + 4
        posEnd   = InStr(posStart, strData, boundary) - 2

        ' Extraer binario
        Dim length
        length = posEnd - posStart
        Dim adoBin
        Set adoBin = CreateObject("ADODB.Stream")
        adoBin.Type = 1
        adoBin.Open
        adoBin.Write binData
        adoBin.Position = posStart - 1
        fileData = adoBin.Read(length)
        adoBin.Close
        Set adoBin = Nothing

        ' Guardar en carpeta uploads
        savePath = Server.MapPath("uploads/" & Year(Now) & Month(Now) & Day(Now) & "_" & fileName)
        SaveBinaryFile savePath, fileData
    End If
End If

' ====================================================
' GUARDAR EN BASE DE DATOS (ejemplo SP)
' ====================================================
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = Conn
    .CommandType = 4 ' adCmdStoredProc
    .CommandText = "usp_Campania_Guardar"

    .Parameters.Append .CreateParameter("@NombreCampania", 200, 1, 200, NombreCampania)
    .Parameters.Append .CreateParameter("@Descripcion", 200, 1, 500, Descripcion)
    .Parameters.Append .CreateParameter("@Canal", 200, 1, 20, Canal)
    .Parameters.Append .CreateParameter("@MensajeWhatsApp", 200, 1, 1000, MensajeWhatsApp)
    .Parameters.Append .CreateParameter("@TemplateWhatsApp", 200, 1, 200, TemplateWhatsApp)
    .Parameters.Append .CreateParameter("@AsuntoEmail", 200, 1, 200, AsuntoEmail)
    .Parameters.Append .CreateParameter("@CuerpoEmail", 200, 1, 2000, CuerpoEmail)
    .Parameters.Append .CreateParameter("@ArchivoRuta", 200, 1, 500, savePath)

    .Execute
End With
Set cmd = Nothing

Response.Write "<p>✅ Campaña guardada correctamente.</p>"
%>
