<%@ Language="VBScript" %>
<!--#include file="includes/FreeASPUpload.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim step
step = Request.QueryString("step")

If step = "" Then step = "1"

' =====================================================
' PASO 1: Mostrar formulario
' =====================================================
If step = "1" Then
%>
    <h2>Paso 1: Subir archivo Excel</h2>
    <form method="post" action="campania_email_configurar.asp?step=2" enctype="multipart/form-data">
        <input type="file" name="archivoExcel">
        <input type="submit" value="Subir archivo">
    </form>
<%
End If

' =====================================================
' PASO 2: Procesar upload
' =====================================================
If step = "2" Then
    Dim Upload, fileName
    Set Upload = New FreeASPUpload
    Upload.Save(Server.MapPath("uploads"))

    If Upload.Form("archivoExcel").FileName <> "" Then
        fileName = Upload.Form("archivoExcel").FileName
        Response.Write "<p>Archivo subido: " & fileName & "</p>"
        ' Guardar el nombre del archivo en sesión para usarlo después
        Session("archivoExcel") = Server.MapPath("uploads/" & fileName)
        Response.Write "<a href='campania_email_configurar.asp?step=3'>Ir al paso 3 (leer Excel)</a>"
    Else
        Response.Write "<p style='color:red'>No se seleccionó archivo.</p>"
    End If
End If

' =====================================================
' PASO 3: Leer el Excel (primer hoja)
' =====================================================
If step = "3" Then
    Dim connExcel, rsExcel, sql, archivo
    archivo = Session("archivoExcel")

    If archivo = "" Then
        Response.Write "<p style='color:red'>No hay archivo cargado.</p>"
    Else
        Set connExcel = Server.CreateObject("ADODB.Connection")
        connExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivo & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

        sql = "SELECT * FROM [Sheet1$]" ' Siempre lee la primer hoja (renombrada o no)

        Set rsExcel = connExcel.Execute(sql)

        Do Until rsExcel.EOF
            Response.Write rsExcel(0).Value & " - " & rsExcel(1).Value & "<br>"
            rsExcel.MoveNext
        Loop

        rsExcel.Close
        connExcel.Close
        Set rsExcel = Nothing
        Set connExcel = Nothing
    End If
End If
%>
