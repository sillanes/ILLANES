<%Class UploadedFile
    Public Name, Data

    Public Sub SetData(fileName, fileBytes)
        Name = fileName
        ' Guardamos los bytes en un Stream para que pueda leerse
        Dim stream
        Set stream = Server.CreateObject("ADODB.Stream")
        stream.Type = 1 ' adTypeBinary
        stream.Open
        stream.Write fileBytes
        stream.Position = 0
        Set Data = stream
    End Sub

    Public Sub SaveTo(path)
        If Not Data Is Nothing Then
            Data.SaveToFile path, 2 ' adSaveCreateOverWrite
            Data.Close
        End If
    End Sub
End Class

Class FreeASPUpload
    Public Files

    Public Sub Class_Initialize()
        Set Files = CreateObject("Scripting.Dictionary")
        Call ParseRequest
    End Sub

    Private Sub ParseRequest()
        Dim bytes(), name, filename, f
        If Request.TotalBytes > 0 Then
            bytes = Request.BinaryRead(Request.TotalBytes)
            
            ' Solo tomamos el archivo con input name "ArchivoExcel"
            name = "ArchivoExcel"
            filename = Request.ServerVariables("HTTP_X_FILE_NAME")
            If filename = "" Then filename = "ArchivoExcel.xlsx"

            Set f = New UploadedFile
            f.SetData filename, bytes
            Files.Add LCase(name), f
        End If
    End Sub
End Class
%>