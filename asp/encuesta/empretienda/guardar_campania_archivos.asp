<%@ Language="VBScript" %>
<%
Option Explicit
Const uploadPath = "C:\Reclamos\Campanias\"

' Crear carpeta si no existe
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(uploadPath) Then fso.CreateFolder(uploadPath)
Set fso = Nothing

' Función para guardar archivo
Function SaveFileSimple(fieldName)
    Dim file, filePath, stream
    On Error Resume Next
    Set file = Request.Files(fieldName)
    If file Is Nothing Then Exit Function
    If file.FileName = "" Then Exit Function

    filePath = uploadPath & file.FileName
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' binario
    stream.Open
    stream.Write file.BinaryRead(file.TotalBytes)
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
    SaveFileSimple = filePath
    On Error GoTo 0
End Function

Response.Write "<h2>Archivos subidos:</h2>"
Response.Write "AdjuntoEmail: " & SaveFileSimple("AdjuntoEmail") & "<br>"
Response.Write "ExcelWhatsApp: " & SaveFileSimple("ExcelWhatsApp") & "<br>"
Response.Write "ExcelEmails: " & SaveFileSimple("ExcelEmails") & "<br>"
%>
