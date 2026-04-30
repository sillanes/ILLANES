<%@ Language="VBScript" %>
<!--#include file="includes/FreeASPUpload.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim upload, file, savePath
Set upload = New FreeASPUpload

Dim CampaniaID
CampaniaID = Request("CampaniaID")
If CampaniaID = "" Then CampaniaID = 1

savePath = Server.MapPath("uploads")

' Buscar archivo subido
Set file = upload.Files("ArchivoExcel")

If file Is Nothing Then
    Response.Write "<p style='color:red'>No se recibió el archivo. Vuelve a intentarlo.</p>"
    Response.End
End If

' Guardar en la carpeta uploads
file.SaveAs savePath & "\" & file.FileName

Response.Write "<p style='color:green'>Archivo subido correctamente: " & file.FileName & "</p>"
%>

<a href="campania_email_configurar2.asp?CampaniaID=<%=CampaniaID%>&step=3">Continuar al paso 3</a>
