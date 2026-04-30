<%@ Language="VBScript" %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim origen, destino, data, fso, streamIn, streamOut, logMsg
origen = "\\192.168.200.14\Scripts\whatsapp_token.txt"
destino = "C:\webhook\whatsapp_token.txt"
logMsg = ""

On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")

' Crear carpeta si no existe
If Not fso.FolderExists("C:\webhook") Then
    fso.CreateFolder("C:\webhook")
End If

If fso.FileExists(origen) Then
    logMsg = logMsg & "✅ Archivo origen encontrado. Intentando leer..." & vbCrLf

    Set streamIn = CreateObject("ADODB.Stream")
    streamIn.Type = 2 ' Texto
    streamIn.Charset = "UTF-8"
    streamIn.Open

    Err.Clear
    streamIn.LoadFromFile origen

    If Err.Number <> 0 Then
        logMsg = logMsg & "❌ Error al leer el archivo de origen: " & Err.Description & vbCrLf
    Else
        data = streamIn.ReadText
        If Len(Trim(data)) = 0 Then
            logMsg = logMsg & "⚠ El archivo se leyó pero está vacío o no se pudo decodificar." & vbCrLf
        Else
            logMsg = logMsg & "✅ Archivo leído correctamente (" & Len(data) & " caracteres)." & vbCrLf
        End If
    End If

    streamIn.Close
    Set streamIn = Nothing
Else
    logMsg = logMsg & "❌ No se encontró el archivo origen: " & origen & vbCrLf
End If

' Guardar destino si hay datos
If Len(Trim(data)) > 0 Then
    Set streamOut = CreateObject("ADODB.Stream")
    streamOut.Type = 2
    streamOut.Charset = "UTF-8"
    streamOut.Open
    streamOut.WriteText data
    streamOut.SaveToFile destino, 2  ' 2 = sobrescribir
    streamOut.Close
    Set streamOut = Nothing
    logMsg = logMsg & "✅ Archivo copiado correctamente a " & destino & vbCrLf
Else
    logMsg = logMsg & "⚠ Archivo destino NO generado con datos (vacío)." & vbCrLf
End If

Set fso = Nothing

Response.Write "<pre>" & logMsg & "</pre>"
%>
