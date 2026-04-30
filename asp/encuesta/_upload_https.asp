<%@ Language=VBScript %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Server.ScriptTimeout = 600

Dim data, filename, bytes, folderPath, filePath
Dim fs, fileOut, decodedData, pos, base64Data
Dim vendedorID, clienteID, objetivoID, periodoID

On Error Resume Next

' === Leer datos ===
data = Request.Form("data")
filename = Request.Form("filename")
clienteID = Request.Form("ClienteID")
objetivoID = Request.Form("ObjetivoID")
periodoID = Request.Form("PeriodoID")

If Session("VendedorID") = "" Then
    vendedorID = "000"
Else
    vendedorID = Trim(Session("VendedorID"))
End If

' Guardar también estas sesiones
Session("ClienteID") = clienteID
Session("ObjetivoID") = objetivoID
Session("PeriodoID") = periodoID

If Len(data) = 0 Or Len(filename) = 0 Then
    Response.Write "❌ No se recibieron datos."
    Response.End
End If

folderPath = "C:\Vendedores\Fotos\" & vendedorID
Set fs = CreateObject("Scripting.FileSystemObject")
If Not fs.FolderExists(folderPath) Then fs.CreateFolder(folderPath)

Dim ext
If InStrRev(filename, ".") > 0 Then
    ext = Mid(filename, InStrRev(filename, "."))
Else
    ext = ".dat"
End If

Dim saveName
saveName = Replace(filename, Chr(34), "")
saveName = Replace(saveName, " ", "_")
filePath = folderPath & "\" & saveName

Function Base64Decode(ByVal strData)
    Dim xml, node
    Set xml = CreateObject("MSXML2.DOMDocument.3.0")
    Set node = xml.createElement("b64")
    node.dataType = "bin.base64"
    node.text = strData
    Base64Decode = node.nodeTypedValue
    Set node = Nothing
    Set xml = Nothing
End Function

decodedData = Base64Decode(data)

Set fileOut = CreateObject("ADODB.Stream")
fileOut.Type = 1
fileOut.Open
fileOut.Write decodedData

If Err.Number <> 0 Then
    Response.Write "❌ Error al guardar archivo: " & Err.Description
    Response.End
End If

fileOut.SaveToFile filePath, 2
fileOut.Close
Set fileOut = Nothing

Session("FileUploaded") = filePath
Session("FileName") = saveName

Dim sizeKB
sizeKB = Round(LenB(decodedData)/1024,1)

Response.Write "✅ Archivo guardado correctamente (" & sizeKB & " KB) como: " & filePath
Response.Flush

Set fs = Nothing
%>
