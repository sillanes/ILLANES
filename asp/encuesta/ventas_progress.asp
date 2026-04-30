<%@ Language="VBScript" %>
<%
Response.Expires = 0
Response.Buffer = True
Response.Clear

Dim UploadID, TempFile
UploadID = Request.QueryString("UploadID")

If UploadID = "" Then
    Response.Write "<html><body><b>Error:</b> Falta UploadID.</body></html>"
    Response.End
End If

Dim fs, tempFolder, tempFileName
Set fs = CreateObject("Scripting.FileSystemObject")

tempFolder = "C:\Vendedores\TempUpload"
tempFileName = tempFolder & "\pu" & UploadID & ".~tmp"

'=== Si no existe el archivo todavía, mostrar mensaje inicial ===
If Not fs.FileExists(tempFileName) Then
    Response.Write "<html><head><meta http-equiv='refresh' content='1'></head><body style='font-family:Arial;text-align:center;margin-top:30px;'>"
    Response.Write "<h3>? Esperando inicio de carga...</h3>"
    Response.Write "</body></html>"
    Response.End
End If

'=== Leer contenido del archivo ===
Dim txt
On Error Resume Next
Dim f : Set f = fs.OpenTextFile(tempFileName, 1)
If Not f Is Nothing Then txt = f.ReadAll
On Error GoTo 0

If LCase(Trim(txt)) = "done" Then
    fs.DeleteFile(tempFileName)
    Response.Write "<html><head><meta http-equiv='refresh' content='1'></head>"
    Response.Write "<body style='font-family:Arial;text-align:center;margin-top:30px;background-color:#f6fff6;'>"
    Response.Write "<h3 style='color:green;'>? Archivo subido correctamente.</h3>"
    Response.Write "<script>setTimeout(function(){window.close();},1500);</script>"
    Response.Write "</body></html>"
    Response.End
End If

Dim parts : parts = Split(txt, vbCrLf)
Dim total, readed, perc
If UBound(parts) >= 2 Then
    total = CLng(parts(1))
    readed = CLng(parts(2))
    If total > 0 Then
        perc = Int((readed / total) * 100)
        If perc > 100 Then perc = 100
    Else
        perc = 0
    End If
Else
    perc = 0
End If

'=== HTML con barra de progreso ===
Response.Write "<html><head>"
Response.Write "<meta http-equiv='refresh' content='1'>"
Response.Write "<title>Progreso de carga</title>"
Response.Write "<style>"
Response.Write "body{font-family:Arial;text-align:center;margin-top:40px;background:#f9f9f9;}"
Response.Write ".bar-container{width:80%;height:25px;background:#ddd;border-radius:20px;margin:0 auto;overflow:hidden;}"
Response.Write ".bar{height:100%;width:" & perc & "%;background:#4CAF50;transition:width 0.5s ease;}"
Response.Write ".percent{margin-top:10px;font-size:14px;font-weight:bold;color:#333;}"
Response.Write "</style>"
Response.Write "</head><body>"
Response.Write "<h3>?? Subiendo archivo...</h3>"
Response.Write "<div class='bar-container'><div class='bar'></div></div>"
Response.Write "<div class='percent'>" & perc & "%</div>"
Response.Write "</body></html>"
%>
