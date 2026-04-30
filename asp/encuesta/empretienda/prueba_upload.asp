<%@ Language="VBScript" %>
<!--#include file="includes/FreeASPUpload.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

' Instanciar parser
Dim Upload
Set Upload = New FreeASPUpload

' Comprobar si llegó archivo (clave en minúsculas)
If Not Upload.Files.Exists("archivoexcel") Then
    Response.Write "<p style='color:red;'>No se recibió el archivo. Verificá el formulario (name='ArchivoExcel') y enctype='multipart/form-data'.</p>"
    Response.End
End If

Dim uf
Set uf = Upload.Files("archivoexcel")

' Crear carpeta uploads si no existe
Dim fso, uploadFolder
Set fso = Server.CreateObject("Scripting.FileSystemObject")
uploadFolder = Server.MapPath("uploads")
If Not fso.FolderExists(uploadFolder) Then
    On Error Resume Next
    fso.CreateFolder(uploadFolder)
    On Error GoTo 0
End If

' Limpiar filename de caracteres inválidos
Dim cleanedName
cleanedName = uf.FileName
cleanedName = Replace(cleanedName, ":", "_")
cleanedName = Replace(cleanedName, "/", "_")
cleanedName = Replace(cleanedName, "\", "_")
cleanedName = Replace(cleanedName, "..", "_")

Dim savePath
savePath = uploadFolder & "\" & cleanedName

' Guardar en disco usando el método SaveTo (gestiona el stream correctamente)
On Error Resume Next
uf.SaveTo savePath
If Err.Number <> 0 Then
    Response.Write "<p style='color:red;'>Error al guardar el archivo: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
    Response.End
End If
On Error GoTo 0

Response.Write "<p style='color:green;'>Archivo subido correctamente: <b>" & Server.HTMLEncode(cleanedName) & "</b></p>"
Response.Write "<p>Guardado en: " & Server.HTMLEncode(savePath) & "</p>"
%>
