<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 60

Dim sql, rs
sql = ""
sql = sql & "SELECT TOP 30 " & vbCrLf
sql = sql & "    h.Numero," & vbCrLf
sql = sql & "    MAX(ISNULL(c.Nombre,'') ) AS Nombre," & vbCrLf
sql = sql & "    ult.Fecha AS UltimaFecha," & vbCrLf
sql = sql & "    LEFT(ult.Contenido, 200) AS UltimoContenido," & vbCrLf
sql = sql & "    CASE WHEN EXISTS (" & vbCrLf
sql = sql & "        SELECT 1 FROM [WhatsAppAPI].sofia.WhatsAPP_API_Contacto cc" & vbCrLf
sql = sql & "        WHERE cc.WhatsAppNumero = h.Numero AND cc.Comentario LIKE '%Chat con Asesor%'" & vbCrLf
sql = sql & "    ) THEN 1 ELSE 0 END AS ChatAsesor" & vbCrLf
sql = sql & "FROM (" & vbCrLf
sql = sql & "    SELECT Numero, MAX(Fecha) AS UltimaFecha" & vbCrLf
sql = sql & "    FROM [WhatsAppAPI].sofia.WhatsAPP_API_Historial" & vbCrLf
sql = sql & "    WHERE Tipo = 'message'" & vbCrLf
sql = sql & "    GROUP BY Numero" & vbCrLf
sql = sql & ") h" & vbCrLf
sql = sql & "CROSS APPLY (" & vbCrLf
sql = sql & "    SELECT TOP 1 Fecha, Contenido" & vbCrLf
sql = sql & "    FROM [WhatsAppAPI].sofia.WhatsAPP_API_Historial h2" & vbCrLf
sql = sql & "    WHERE h2.Numero = h.Numero AND h2.Tipo = 'message'" & vbCrLf
sql = sql & "    ORDER BY h2.Fecha DESC, h2.ID DESC" & vbCrLf
sql = sql & ") ult" & vbCrLf
sql = sql & "LEFT JOIN [WhatsAppAPI].sofia.WhatsAPP_API_Contacto c" & vbCrLf
sql = sql & "    ON c.WhatsAppNumero = h.Numero" & vbCrLf
sql = sql & "GROUP BY H.Numero, ult.Fecha , LEFT(ult.Contenido, 200) " & vbCrLf
sql = sql & "ORDER BY ult.Fecha DESC;" & vbCrLf

'response.write "<pre>" & sql & "</pre>"
Set rs = conn.Execute(sql)
%>
<html>
<head>
<meta charset="utf-8">
<style>
body{margin:0;font-family:'Segoe UI',sans-serif;}
.chat-item{padding:10px 12px;border-bottom:1px solid #eee;cursor:pointer;display:flex;flex-direction:column;gap:4px}
.chat-item:hover{background:#f8f8f8;}
.row1{display:flex;justify-content:space-between;align-items:center}
.chat-numero{font-weight:bold;color:#333;}
.chat-nombre{color:#555;font-size:13px;}
.chat-prev{font-size:12px;color:#777;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.chat-fecha{font-size:11px;color:#aaa;}
.chat-asesor{color:#d39e00;font-weight:bold;margin-left:6px;}
</style>
</head>
<body>
<%
If rs.EOF Then
    Response.Write "<div style='padding:20px;color:#777;'>No hay conversaciones registradas.</div>"
Else
    Do Until rs.EOF
        Dim numero, nombre, ultMensaje, fecha, asesorFlag
        numero     = rs("Numero")
        nombre     = rs("Nombre")
        ultMensaje = rs("UltimoContenido")
        fecha      = rs("UltimaFecha")
        asesorFlag = rs("ChatAsesor")

        If IsNull(ultMensaje) Or ultMensaje = "" Then ultMensaje = "(sin texto)"
        If IsNull(nombre) Then nombre = ""

        Response.Write "<div class='chat-item' onclick=""parent.location='mensajes_sofia.asp?numero=" & numero & "';"">"
        Response.Write   "<div class='row1'>"
        Response.Write     "<div class='chat-numero'>" & Server.HTMLEncode(numero) & "</div>"
        Response.Write     "<div class='chat-fecha'>"  & fecha & "</div>"
        Response.Write   "</div>"
        If nombre <> "" Then
            Response.Write "<div class='chat-nombre'>" & Server.HTMLEncode(nombre) & "</div>"
        End If
        Response.Write   "<div class='chat-prev'>" & Server.HTMLEncode(ultMensaje) & "</div>"
        If asesorFlag = 1 Then
            Response.Write "<div class='chat-asesor'>🗣️ Chat con asesor</div>"
        End If
        Response.Write "</div>"

        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
%>
</body>
</html>
