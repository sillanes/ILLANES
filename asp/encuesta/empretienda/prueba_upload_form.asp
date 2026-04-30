<%@ Language="VBScript" %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim msg
msg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim binData, fileName, fso, savePath
    binData = Request.BinaryRead(Request.TotalBytes)

    ' Solo para este ejemplo, nombre fijo:
    fileName = "ArchivoExcel_" & Replace(Replace(Now, ":", "_"), "/", "_") & ".xlsx"
    
    savePath = Server.MapPath("uploads") & "\" & fileName

    ' Crear objeto FSO
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Server.MapPath("uploads")) Then
        fso.CreateFolder(Server.MapPath("uploads"))
    End If

    ' Crear archivo binario y escribir
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write binData
    stream.SaveToFile savePath, 2 ' 2 = sobrescribir si existe
    stream.Close
    Set stream = Nothing
    Set fso = Nothing

    msg = "<p style='color:green;'>Archivo subido correctamente: <b>" & fileName & "</b></p>"
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Subida de Archivo en ASP Clásico</title>
</head>
<body>
<h2>Subir archivo</h2>
<form method="POST" enctype="multipart/form-data">
    <input type="file" name="ArchivoExcel" required>
    <button type="submit">Subir Archivo</button>
</form>
<div><%=msg%></div>
</body>
</html>
