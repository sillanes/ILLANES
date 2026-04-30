<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.ContentType = "application/json"

Dim numero
numero = Trim(Request.QueryString("numero"))
If numero = "" Then
    Response.Write "[]"
    Response.End
End If

' ==========================
' CONSULTA SQL
' ==========================
Dim sql
sql = "SELECT Estado, Contenido, " & _
      "CONVERT(VARCHAR(19), Fecha, 120) AS FechaSQL, " & _
      "ISNULL(Origen,'usuario') AS Origen " & _
      "FROM [WhatsAppAPI].sofia.WhatsAPP_API_Historial " & _
      "WHERE Numero = '" & Replace(numero,"'","") & "' " & _
      "AND Tipo='message' " & _
      "ORDER BY Fecha ASC"

' ==========================
' ABRIR RECORDSET MANUALMENTE
' ==========================
Dim rs : Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3  ' adUseClient
rs.Open sql, conn, 1, 1  ' adOpenKeyset, adLockReadOnly

Function EscapeJSON(str)
    If IsNull(str) Then
        EscapeJSON = ""
        Exit Function
    End If
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCrLf, "\n")
    str = Replace(str, vbCr, "\n")
    str = Replace(str, vbLf, "\n")
    EscapeJSON = str
End Function

' ==========================
' CONSTRUIR JSON
' ==========================
Dim mensajes : mensajes = "["
Do Until rs.EOF
    Dim contenido, estado, origen, fecha
    contenido = Trim(rs("Contenido") & "")
    estado    = Trim(rs("Estado") & "")
    origen    = Trim(rs("Origen") & "")
    fecha     = Trim(rs("FechaSQL") & "")

    ' Si SQL devolvió fecha, formatear a ISO
    If Len(fecha) > 0 Then fecha = Replace(fecha, " ", "T")

    mensajes = mensajes & "{""Estado"":""" & EscapeJSON(estado) & """," & _
                            """Contenido"":""" & EscapeJSON(contenido) & """," & _
                            """Fecha"":""" & EscapeJSON(fecha) & """," & _
                            """Origen"":""" & EscapeJSON(origen) & """}"
    rs.MoveNext
    If Not rs.EOF Then mensajes = mensajes & ","
Loop
mensajes = mensajes & "]"

rs.Close : Set rs = Nothing

Response.Write mensajes
%>
