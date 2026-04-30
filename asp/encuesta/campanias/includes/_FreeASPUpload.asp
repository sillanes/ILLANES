<%
' FreeASPUpload.asp - versión funcional mínima
Class FreeASPUpload
    Private mFileData, mFileName

    Public Property Get FileData()
        Set FileData = mFileData
    End Property

    Public Property Get FileName()
        FileName = mFileName
    End Property

    Public Sub LoadFile(formField)
        Dim binData, totalBytes
        totalBytes = Request.TotalBytes
        If totalBytes = 0 Then
            Response.Write "No se recibió el archivo."
            Response.End
        End If

        binData = Request.BinaryRead(totalBytes)

        ' Para simplificar, guardamos directamente el archivo
        Dim fso, filePath
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        mFileName = formField ' el nombre real del archivo se puede obtener de los headers si quieres
        filePath = Server.MapPath("uploads") & "\" & mFileName
        Dim stream
        Set stream = Server.CreateObject("ADODB.Stream")
        stream.Type = 1 ' binary
        stream.Open
        stream.Write binData
        stream.SaveToFile filePath, 2 ' 2 = sobrescribir
        stream.Close
        Set stream = Nothing
        Set fso = Nothing
    End Sub
End Class
%>
