<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="_upload.asp" -->
<%
If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

Dim hdrid, clienteid, doWhat
hdrid     = Request.QueryString("hdrid")
clienteid = Request.QueryString("clienteid")
doWhat    = Request.QueryString("doWhat")

If hdrid = "" Or clienteid = "" Then
    Response.Write "Datos incompletos"
    Response.End
End If

Server.ScriptTimeout = 3600

' =========================
' CONFIG UPLOAD
' =========================
Dim Form : Set Form = New ASPForm
Form.UploadID  = Request.QueryString("UploadID")
Form.SizeLimit = 10 * 1024 * 1024 ' 10 MB

Const fsCompletted = 0
Const fsError      = &HA
Const fsSizeLimit  = &HD
Const fsTimeOut    = &HE

Dim msg
msg = ""

If Form.State > fsError Then
    If Form.State = fsSizeLimit Then
        msg = "El archivo supera el tamaño permitido."
    ElseIf Form.State = fsTimeOut Then
        msg = "Tiempo de carga excedido."
    Else
        msg = "Error de carga (estado " & Form.State & ")."
    End If

    Response.Write msg
    Response.End
End If

' =========================
' CUANDO EL UPLOAD FINALIZA
' =========================
If Form.State = fsCompletted Then

    ' Guardar archivo (ruta RELATIVA como usa tu sistema)
    Dim relFolder
    relFolder = "comprobantes/"

    Form.Files.Save MapFolderToDisk(relFolder)

    Dim f, fileName, ext
    Set f = Form.Files(1)

    fileName = f.FileName
    ext = LCase(Mid(fileName, InStrRev(fileName,".") + 1))

    If ext <> "pdf" And ext <> "jpg" And ext <> "jpeg" And ext <> "png" Then
        Response.Write "Tipo de archivo no permitido"
        Response.End
    End If

    ' Guardar en sesión (MISMO MECANISMO QUE CAMPAÑAS)
    Session("FileUploaded") = fileName

    ' Guardar referencia en SQL
    Dim sql
    sql = "EXEC cobranza.Transportista_Comprobante_Ins " & _
          hdrid & "," & clienteid & "," & _
          "'" & Replace(fileName,"'","''") & "'," & _
          "'" & Replace(relFolder,"'","''") & "'," & _
          "'" & ext & "'," & _
          "'" & Replace(Session("NombreTransportista"),"'","''") & "'"

    conn.Execute sql

    ' Volver a la hoja de ruta
    Response.Redirect "hojaderutav3.asp?hdrid=" & hdrid & "&msg=upload_ok"
End If
%>

<!-- ========================= -->
<!-- FORMULARIO (IGUAL AL STEP 4) -->
<!-- ========================= -->

<form method="post"
      enctype="multipart/form-data"
      action="hojaderuta_upload.asp?hdrid=<%=hdrid%>&clienteid=<%=clienteid%>&doWhat=3&UploadID=<%=Form.NewUploadID%>">

    <input type="file" name="File1" required>

    <button type="submit">Subir archivo</button>
</form>
