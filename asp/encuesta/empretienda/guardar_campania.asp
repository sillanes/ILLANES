<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<% 

Const uploadPath = "C:\Reclamos\Campanias\"
Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(uploadPath) Then fso.CreateFolder(uploadPath)
Set fso = Nothing

' ---------- Leer datos ----------
Dim NombreCampania, Descripcion, Canal
Dim MensajeWhatsApp, TemplateWhatsApp
Dim MensajeEmailAsunto, MensajeEmailCuerpo
NombreCampania = Request.Form("NombreCampania")
Descripcion = Request.Form("Descripcion")
Canal = Request.Form("Canal")
MensajeWhatsApp = Request.Form("MensajeWhatsApp")
TemplateWhatsApp = Request.Form("TemplateWhatsApp")
MensajeEmailAsunto = Request.Form("MensajeEmailAsunto")
MensajeEmailCuerpo = Request.Form("MensajeEmailCuerpo")

response.write "es joda"
response.write MensajeWhatsApp
' ---------- Guardar en DB ----------
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = Conn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "usp_Campania_Insertar"
cmd.Parameters.Append cmd.CreateParameter("@NombreCampania", 200, 1, 200, NombreCampania)
cmd.Parameters.Append cmd.CreateParameter("@Descripcion", 200, 1, 500, Descripcion)
cmd.Parameters.Append cmd.CreateParameter("@Canal", 200, 1, 50, Canal)
cmd.Parameters.Append cmd.CreateParameter("@MensajeWhatsApp", 200, 1, 4000, MensajeWhatsApp)
cmd.Parameters.Append cmd.CreateParameter("@TemplateWhatsApp", 200, 1, 200, TemplateWhatsApp)
cmd.Parameters.Append cmd.CreateParameter("@MensajeEmailAsunto", 200, 1, 200, MensajeEmailAsunto)
cmd.Parameters.Append cmd.CreateParameter("@MensajeEmailCuerpo", 200, 1, 4000, MensajeEmailCuerpo)
Set rs = cmd.Execute()
Dim CampaniaID : CampaniaID = rs(0)
Set cmd = Nothing

' ---------- Guardar archivos ----------
Function SaveFile(fieldName)
    Dim file, filePath, stream
    On Error Resume Next
    Set file = Request.Files(fieldName)
    If Not file Is Nothing Then
        If file.FileName <> "" Then
            filePath = uploadPath & file.FileName
            Set stream = Server.CreateObject("ADODB.Stream")
            stream.Type = 1
            stream.Open
            stream.Write file.BinaryRead(file.TotalBytes)
            stream.SaveToFile filePath, 2
            stream.Close
            Set stream = Nothing

            ' Guardar ruta en DB
            Dim cmdFile
            Set cmdFile = Server.CreateObject("ADODB.Command")
            cmdFile.ActiveConnection = Conn
            cmdFile.CommandType = 4
            cmdFile.CommandText = "usp_Campania_Archivo_Insertar"
            cmdFile.Parameters.Append cmdFile.CreateParameter("@CampaniaID", 3, 1, , CampaniaID)
            cmdFile.Parameters.Append cmdFile.CreateParameter("@TipoArchivo", 200, 1, 50, fieldName)
            cmdFile.Parameters.Append cmdFile.CreateParameter("@RutaArchivo", 200, 1, 500, filePath)
            cmdFile.Execute
            Set cmdFile = Nothing
        End If
    End If
    On Error GoTo 0
End Function

SaveFile "AdjuntoEmail"
SaveFile "ExcelWhatsApp"
SaveFile "ExcelEmails"

Response.Write "<h2>Campaña y archivos guardados correctamente.</h2>"
Response.Write "<a href='campania_form.asp'>Crear otra campaña</a>"
%>
