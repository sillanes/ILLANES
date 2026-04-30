<%@ Language="VBScript" %>
<%
Response.Buffer = True
Response.Clear

Dim fileParam, filePath, fso, fileName, pos

fileParam = Request.QueryString("file") ' Ej: A00090055978101FA0000031806.PDF

If fileParam = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "Archivo no especificado"
    Response.End
End If

' Ruta local
filePath = "C:\Administracion\PDF\" & Replace(fileParam, "/", "\")

Set fso = Server.CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(filePath) Then
    Response.Status = "404 Not Found"
    Response.Write "Archivo no encontrado: " & filePath
    Response.End
End If

' Nombre del archivo
pos = InStrRev(filePath, "\")
If pos > 0 Then
    fileName = Mid(filePath, pos + 1)
Else
    fileName = filePath
End If

' Enviar PDF
Response.ContentType = "application/pdf"
Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
Response.BinaryWrite ReadBinaryFile(filePath)
Response.Flush
Response.End

Set fso = Nothing

' =======================================
' Función para leer archivo binario
Function ReadBinaryFile(filePath)
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    ReadBinaryFile = stream.Read
    stream.Close
    Set stream = Nothing
End Function
%>
