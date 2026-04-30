<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="includes/FreeASPUpload.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim upload, f, filename, filedata, rs
Set upload = New FreeASPUpload

If upload.Files.Count = 0 Then
    Response.Write "<p style='color:red;'>No se recibió el archivo. Vuelve a intentarlo.</p>"
    Response.End
End If

For Each f In upload.Files
    filename = upload.Files(f)(0)
    filedata = upload.Files(f)(1)

    ' Convertir a stream para leer Excel
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write filedata
    stream.Position = 0

    ' Abrir conexión a Excel en memoria (solo hoja 1)
    Dim cnExcel, rsExcel, sql
    Set cnExcel = Server.CreateObject("ADODB.Connection")
    cnExcel.Provider = "Microsoft.ACE.OLEDB.12.0"
    cnExcel.Open "Data Source=" & filename & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    Set rsExcel = cnExcel.Execute("SELECT * FROM [Hoja1$]") ' siempre la primer hoja

    ' Insertar en DB
    Do Until rsExcel.EOF
        sql = "INSERT INTO dbo.Campania_Emails (CampaniaID, Email, Nombre) VALUES (" & _
            Request.QueryString("CampaniaID") & ", '" & Replace(rsExcel("Email"), "'", "''") & "', '" & Replace(rsExcel("Nombre"), "'", "''") & "')"
        conn.Execute sql
        rsExcel.MoveNext
    Loop

    rsExcel.Close
    cnExcel.Close
    Set rsExcel = Nothing
    Set cnExcel = Nothing
    stream.Close
    Set stream = Nothing
Next

Response.Write "<p style='color:green;'>Archivo procesado e insertado correctamente ✅</p>"
%>
