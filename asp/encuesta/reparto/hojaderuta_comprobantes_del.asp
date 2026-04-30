<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
If Session("NombreTransportista") = "" Then
  Response.Redirect "login.asp"
End If

Dim id, hdrid, clienteid
id        = Request("id")
hdrid     = Request("hdrid")
clienteid = Request("clienteid")

If id = "" Then
  Response.End
End If

Dim rs, fullPath

' 1️⃣ Obtener datos del comprobante vía SP
Set rs = conn.Execute("EXEC cobranza.Transportista_Comprobante_Get " & CLng(id))

If Not rs.EOF Then

    fullPath = rs("RutaArchivo") &"\"& rs("NombreArchivo")

    rs.Close
    Set rs = Nothing

    ' 2️⃣ Borrar archivo físico
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(fullPath) Then
        fso.DeleteFile fullPath, True
    End If

    Set fso = Nothing

    ' 3️⃣ Borrar registro vía SP
    conn.Execute "EXEC cobranza.Transportista_Comprobante_Del " & CLng(id)

End If

conn.Close
Set conn = Nothing

' 4️⃣ Volver a la grilla
Response.Redirect "hojaderuta_comprobantes.asp?hdrid=" & hdrid & "&clienteid=" & clienteid
%>
