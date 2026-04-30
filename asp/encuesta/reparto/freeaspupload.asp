<%
Class UploadedFile
  Public ContentType, FileName, Size, Data
End Class

Class FreeASPUpload
  Public UploadedFiles, Form
  Private BoundaryB

  Private Sub Class_Initialize()
    Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
    Set Form = Server.CreateObject("Scripting.Dictionary")
  End Sub

  Public Sub Upload()
    Dim binData, pos, part, headerEnd, dataStart, dataEnd
    Dim headerStr, name, fileName, contentType
    Dim ct, boundaryStr

    ' ===== LEER REQUEST =====
    binData = Request.BinaryRead(Request.TotalBytes)

    ' ===== BOUNDARY DESDE HEADER HTTP =====
    ct = Request.ServerVariables("CONTENT_TYPE")
    If InStr(ct, "boundary=") = 0 Then Exit Sub

    boundaryStr = "--" & Mid(ct, InStr(ct, "boundary=") + 9)

    ' ===== CONVERTIR BOUNDARY A BYTES (CLAVE) =====
    BoundaryB = StringToBytes(boundaryStr)

    pos = 1

    Do
      part = InStrB(pos, binData, BoundaryB)
      If part = 0 Then Exit Do

      headerEnd = InStrB(part, binData, ChrB(13) & ChrB(10) & ChrB(13) & ChrB(10))
      If headerEnd = 0 Then Exit Do

      headerStr = CStr(MidB(binData, part, headerEnd - part))

      name = ""
      If InStr(headerStr, "name=""") > 0 Then
        name = Mid(headerStr, InStr(headerStr, "name=""") + 6)
        name = Left(name, InStr(name, """") - 1)
      End If

      If InStr(headerStr, "filename=""") > 0 Then
        fileName = Mid(headerStr, InStr(headerStr, "filename=""") + 10)
        fileName = Left(fileName, InStr(fileName, """") - 1)

        contentType = ""
        If InStr(headerStr, "Content-Type:") > 0 Then
          contentType = Trim(Mid(headerStr, InStr(headerStr, "Content-Type:") + 13))
        End If

        dataStart = headerEnd + 4
        dataEnd = InStrB(dataStart, binData, BoundaryB) - 2
        If dataEnd <= dataStart Then Exit Do

        Dim uf
        Set uf = New UploadedFile
        uf.FileName = fileName
        uf.ContentType = contentType
        uf.Data = MidB(binData, dataStart, dataEnd - dataStart)
        uf.Size = LenB(uf.Data)

        UploadedFiles.Add UploadedFiles.Count + 1, uf
      End If

      pos = dataEnd + 2
    Loop
  End Sub
End Class

' ===== UTIL: STRING → BYTES =====
Function StringToBytes(str)
  Dim i, b
  b = ""
  For i = 1 To Len(str)
    b = b & ChrB(Asc(Mid(str, i, 1)))
  Next
  StringToBytes = b
End Function
%>
