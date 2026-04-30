<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
On Error Resume Next

Dim tipo, cmd, rs, mes, anio
tipo = LCase(Trim(Request("tipo")))

mes = Month(Date())
anio = Year(Date())


Set cmd = Server.CreateObject("ADODB.Command")
With cmd
	.ActiveConnection = conn
	.CommandType = 4
	.CommandText = "reclamos.[report].[usp_Pxp_Eficiencia_Periodo_Actual_Sel]"
	Set rs = .Execute()
End With

If Not rs.EOF Then periodoNombre = rs("PeriodoCodigo") & "/" & rs("PeriodoAnio")



' =====================================================================================
' REPORTE 1: EFICIENCIA DE ARMADORES (CON TÃTULO)
' =====================================================================================
If tipo = "eficiencia" Then

    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandType = 4
        .CommandText = "reclamos.[report].[usp_Pxp_Eficiencia_Armadores_PeriodoActual]"
        Set rs = .Execute()
    End With
 

    ' ===== TÃTULO =====
    Response.Write "<h3 style='margin:0 0 10px 0; color:#0059B3; font-weight:bold;'>"
    If periodoNombre <> "" Then
        Response.Write "Eficiencia de Armadores - " & periodoNombre
    Else
        Response.Write "Eficiencia de Armadores"
    End If
    Response.Write "</h3>"

    ' ===== TABLA =====
    Response.Write "<table class='sortable'>"
    Response.Write "<thead><tr>"
    Response.Write "<th class='sortable-th'>Nombre</th>"
    Response.Write "<th class='sortable-th'>Cant. Pedidos</th>"
    Response.Write "<th class='sortable-th'>Cant. Errores</th>"
    Response.Write "<th class='sortable-th'>Eficiencia</th>"
    Response.Write "</tr></thead><tbody>"

    If rs.EOF Then
        Response.Write "<tr><td colspan='4'>Sin datos</td></tr>"
    Else

        rs.MoveFirst ' <<< importante: volver al primer registro

        Do Until rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("Nombre") & "</td>"
            Response.Write "<td style='text-align:right;'>" & rs("CantPedidos") & "</td>"
            Response.Write "<td style='text-align:right;'>" & rs("CantErrores") & "</td>"
            Response.Write "<td style='text-align:right;'>" & rs("Eficiencia") & "%</td>"
            Response.Write "</tr>"
            rs.MoveNext
        Loop
    End If

    Response.Write "</tbody></table>"
    Response.End
End If


' =====================================================================================
' REPORTE 2: ERRORES POR ARMADOR (CON TÃTULO)
' =====================================================================================
If tipo = "errores" Then

    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandType = 4
        .CommandText = "reclamos.report.[usp_Pxp_Errores_Dia]"
        Set rs = .Execute()
    End With
 

    ' ===== TITULO =====
    Response.Write "<h3 style='margin:0 0 10px 0; color:#B30000; font-weight:bold;'>"
    If periodoNombre <> "" Then
        Response.Write "Errores por Armador - " & periodoNombre
    Else
        Response.Write "Errores por Armador"
    End If
    Response.Write "</h3>"

    ' ===== TABLA =====
    Response.Write "<table class='sortable'>"
    Response.Write "<thead><tr>"
    Response.Write "<th class='sortable-th'>Nombre</th>"
    Response.Write "<th class='sortable-th'>Cant. Errores</th>"
    Response.Write "<th class='sortable-th'>Detalle</th>"
    Response.Write "</tr></thead><tbody>"

    If rs.EOF Then
        Response.Write "<tr><td colspan='3'>Sin datos</td></tr>"
    Else

        rs.MoveFirst

        Do Until rs.EOF

            Dim detalleRaw, partes, detalleFmt, i
            detalleRaw = Server.HTMLEncode(rs("Detalle"))
            partes = Split(detalleRaw, "-")

            detalleFmt = ""
            For i = 0 To UBound(partes)
                Dim linea
                linea = Trim(partes(i))
                If linea <> "" Then detalleFmt = detalleFmt & " " & linea & vbCrLf
            Next

            Response.Write "<tr>"
            Response.Write "<td>" & rs("Nombre") & "</td>"
            Response.Write "<td style='text-align:right;'>" & rs("CantidadErrores") & "</td>"
            Response.Write "<td class='detalle-cell'>" & detalleFmt & "</td>"
            Response.Write "</tr>"

            rs.MoveNext
        Loop
    End If

    Response.Write "</tbody></table>"
    Response.End
End If
%>
