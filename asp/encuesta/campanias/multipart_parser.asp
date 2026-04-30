<%
Class MultipartParser
    Public Fields
    Public Files
    Private m_boundary
    Private m_savePath

    Private Sub Class_Initialize()
        Set Fields = Server.CreateObject("Scripting.Dictionary")
        Set Files = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub ParseRequest(savePath)
        Dim binData, parts, i, part, name, filename, value, posCRLF, fileData
        Dim fso, fileStream

        ' Guardar path de destino
        m_savePath = savePath

        ' Verificar que la carpeta exista
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(m_savePath) Then
            fso.CreateFolder(m_savePath)
        End If
        Set fso = Nothing

        ' Leer todo el request como binario
        binData = Request.BinaryRead(Request.TotalBytes)
        binData = BinaryToString(binData)

        ' Obtener boundary
        m_boundary = Left(binData, InStr(binData, vbCrLf) - 1)

        ' Separar partes
        parts = Split(binData, m_boundary)

        For i = 0 To UBound(parts)
            part = parts(i)

            If InStr(part, "Content-Disposition") > 0 Then
                ' extraer nombre
                name = ExtractBetween(part, "name=""", """")

                ' revisar si es archivo o campo de texto
                If InStr(part, "filename=""") > 0 Then
                    filename = ExtractBetween(part, "filename=""", """")
                    If filename <> "" Then
                        ' obtener binario del archivo
                        posCRLF = InStr(part, vbCrLf & vbCrLf)
                        If posCRLF > 0 Then
                            fileData = MidB(part, posCRLF + 4, LenB(part) - (posCRLF + 6))

                            ' guardar archivo
                            Set fileStream = CreateObject("ADODB.Stream")
                            fileStream.Type = 1 'binary
                            fileStream.Open
                            fileStream.Write fileData
                            fileStream.SaveToFile m_savePath & "\" & GetFileName(filename), 2
                            fileStream.Close
                            Set fileStream = Nothing

                            ' registrar archivo
                            Files.Add name, m_savePath & "\" & GetFileName(filename)
                        End If
                    End If
                Else
                    ' es un campo de texto
                    posCRLF = InStr(part, vbCrLf & vbCrLf)
                    If posCRLF > 0 Then
                        value = Trim(Mid(part, posCRLF + 4))
                        value = Left(value, Len(value) - 2) ' quitar CRLF final
                        Fields.Add name, value
                    End If
                End If
            End If
        Next
    End Sub

    ' =====================
    ' Helpers
    ' =====================
    Private Function ExtractBetween(ByVal src, ByVal a, ByVal b)
        Dim p1, p2
        p1 = InStr(src, a)
        If p1 > 0 Then
            p1 = p1 + Len(a)
            p2 = InStr(p1, src, b)
            If p2 > 0 Then
                ExtractBetween = Mid(src, p1, p2 - p1)
            Else
                ExtractBetween = ""
            End If
        Else
            ExtractBetween = ""
        End If
    End Function

    Private Function GetFileName(fullPath)
        Dim parts
        parts = Split(fullPath, "\")
        GetFileName = parts(UBound(parts))
    End Function

    Private Function BinaryToString(bin)
        Dim stm
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 1
        stm.Open
        stm.Write bin
        stm.Position = 0
        stm.Type = 2
        stm.Charset = "ISO-8859-1"
        BinaryToString = stm.ReadText
        stm.Close
        Set stm = Nothing
    End Function
End Class
%>
