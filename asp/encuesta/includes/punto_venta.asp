<%
Dim host, PUNTO_VENTA_ID
host = LCase(Trim(Request.ServerVariables("HTTP_HOST")))

If InStr(host, ":") > 0 Then host = Split(host, ":")(0) 'por si viene con puerto

Select Case host
    Case "illanes-encuesta.ddns.net"
        PUNTO_VENTA_ID = 1
    Case "illanes-oeste.ddns.net"
        PUNTO_VENTA_ID = 2
    Case Else
        Response.Write "Host no reconocido: " & Server.HTMLEncode(host)
        Response.End
End Select

Session("PUNTO_VENTA_ID") = PUNTO_VENTA_ID
%>
