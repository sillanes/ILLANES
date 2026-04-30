<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 3600

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Function Nz(v, defaultValue)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Function HtmlExcel(v)
    HtmlExcel = Server.HTMLEncode("" & Nz(v, ""))
End Function

Function FormatDateAR(v)
    If IsDate(v) Then
        FormatDateAR = Right("0" & Day(v),2) & "/" & Right("0" & Month(v),2) & "/" & Year(v)
    Else
        FormatDateAR = ""
    End If
End Function

Function FieldExists(rs, fieldName)
    Dim f
    FieldExists = False

    If Not IsObject(rs) Then Exit Function

    For Each f In rs.Fields
        If LCase(f.Name) = LCase(fieldName) Then
            FieldExists = True
            Exit Function
        End If
    Next
End Function

Function GetFieldValue(rs, fieldName, defaultValue)
    If FieldExists(rs, fieldName) Then
        GetFieldValue = Nz(rs(fieldName), defaultValue)
    Else
        GetFieldValue = defaultValue
    End If
End Function

Function BoolSiNo(v)
    Dim s
    s = UCase(Trim("" & Nz(v, "")))

    If s = "1" Or s = "TRUE" Or s = "SI" Or s = "S" Then
        BoolSiNo = "SI"
    Else
        BoolSiNo = "NO"
    End If
End Function

Dim camionID, sql, rs, nombreArchivo
camionID = CLng("0" & Request("CamionID"))

If camionID <= 0 Then
    Response.Write "CamionID inválido"
    Response.End
End If

sql = "EXEC dbo.[usp_Camiones_Items_Exportar] " & _
      "@CamionID=" & camionID   

Set rs = conn.Execute(sql)

nombreArchivo = "camion_" & camionID & "_items_control.xls"

Response.Clear
Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & nombreArchivo

Response.Write "<html lang='es'>"
Response.Write "<head>"
Response.Write "<meta charset='utf-8'>"
Response.Write "<style>"
Response.Write "table{border-collapse:collapse;font-family:Arial;font-size:12px;}"
Response.Write "td,th{border:1px solid #999;padding:6px;}"
Response.Write "th{background:#e9ecef;font-weight:bold;}"
Response.Write ".titulo{font-size:18px;font-weight:bold;margin-bottom:10px;}"
Response.Write ".sub{font-family:Arial;font-size:13px;margin-bottom:12px;}"
Response.Write "</style>"
Response.Write "</head>"
Response.Write "<body>"

Response.Write "<div class='titulo'>Control de Camión #" & camionID & "</div>"

If Not rs.EOF Then
    Response.Write "<div class='sub'>"
    Response.Write "<strong>Proveedor:</strong> " & HtmlExcel(GetFieldValue(rs, "ProveedorNombre", "")) & "<br>"
    Response.Write "<strong>Factura:</strong> " & HtmlExcel(GetFieldValue(rs, "NumeroFactura", "")) & "<br>"
    Response.Write "<strong>Remito:</strong> " & HtmlExcel(GetFieldValue(rs, "NroRemito", "")) & "<br>"
    Response.Write "<strong>Orden de Compra:</strong> " & HtmlExcel(GetFieldValue(rs, "OrdenCompra", "")) & "<br>"
    Response.Write "<strong>Chofer:</strong> " & HtmlExcel(GetFieldValue(rs, "Chofer", "")) & "<br>"
    Response.Write "<strong>Nro Carga:</strong> " & HtmlExcel(GetFieldValue(rs, "NroCarga", "")) & "<br>"
    Response.Write "</div><br>"
End If

Response.Write "<table>"
Response.Write "<tr>"
Response.Write "<th>Factura</th>"
Response.Write "<th>Proveedor</th>"
Response.Write "<th>Fecha Factura</th>"
Response.Write "<th>Remito</th>"
Response.Write "<th>Orden Compra</th>"
Response.Write "<th>ItemID</th>"
Response.Write "<th>Artículo</th>"
Response.Write "<th>Descripción</th>"
Response.Write "<th>Unidad</th>"
Response.Write "<th>Cantidad Sistema</th>"
Response.Write "<th>Cantidad Controlada</th>"
Response.Write "<th>Completo</th>"
Response.Write "<th>Observación</th>"
Response.Write "</tr>"

If Not rs.EOF Then
    Do Until rs.EOF
        Response.Write "<tr>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "NumeroFactura", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "ProveedorNombre", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(FormatDateAR(GetFieldValue(rs, "FechaFactura", ""))) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "NroRemito", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "OrdenCompra", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "ItemID", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "Articulo", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "Descripcion", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "UnidadMedida", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "Cantidad", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "CantidadControlada", "")) & "</td>"
        Response.Write "<td>" & HtmlExcel(BoolSiNo(GetFieldValue(rs, "Completo", 0))) & "</td>"
        Response.Write "<td>" & HtmlExcel(GetFieldValue(rs, "Observacion", "")) & "</td>"
        Response.Write "</tr>"

        rs.MoveNext
    Loop
End If

Response.Write "</table>"
Response.Write "</body>"
Response.Write "</html>"

If IsObject(rs) Then
    On Error Resume Next
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
End If

conn.Close
Set conn = Nothing
%>