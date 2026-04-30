<%@ Language="VBScript" %>
<!--#include file="includes/FreeASPUpload.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim upload, file, savePath, CampaniaID
CampaniaID = Request("CampaniaID")

Set upload = New FreeASPUpload

If upload.Files.Count = 0 Then
    Response.Write "<p style='color:red'>No se recibió el archivo. Vuelve e inténtalo de nuevo.</p>"
    Response.End
End If

' Carpeta donde se guardan los archivos
savePath = Server.MapPath("uploads")

' Asegurar que la carpeta exista
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(savePath) Then
    fso.CreateFolder(savePath)
End If
Set fso = Nothing

' Guardar archivo
Set file = upload.Files("ArchivoExcel")
file.SaveToDisk savePath

Response.Write "<p style='color:green'>Archivo subido correctamente: " & file.FileName & "</p>"

' TODO: Aquí agregamos lógica para leer Excel y registrar en Campania_Destinatario
%>
<a href="campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=3">Continuar al paso 3</a>
