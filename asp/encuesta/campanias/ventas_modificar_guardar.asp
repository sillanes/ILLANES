<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If


Dim orderBy, orderDir, allowedColumns
orderBy = Request("order")
orderDir = UCase(Request("dir"))


response.write "order=" & orderBy

' Columnas permitidas
allowedColumns = "Cliente,VendedorID,Zona,OrdenID,Monto,Pagado"

If orderBy = "" Or InStr(1, allowedColumns, orderBy, vbTextCompare) = 0 Then
    orderBy = "Cliente" ' default
End If

If orderDir <> "DESC" Then
    orderDir = "ASC"
End If

Dim FechaInicio, FechaFin
If Request("FechaInicio") <> "" Then
    FechaInicio = CDate(Request("FechaInicio"))
Else
    FechaInicio = DateAdd("m",-1,Date()) ' default 1 mes atrás
End If

If Request("FechaFin") <> "" Then
    FechaFin = CDate(Request("FechaFin"))
Else
    FechaFin = Date()
End If

Function toISO(fecha)
    Dim d
    d = CDate(fecha)
    toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function

Dim OrdenID, FileID, VendedorID, FechaEntrega, Zona, ObservacionesAdicional
FileID = Request("FileID")
OrdenID = Request("OrdenID")
VendedorID = Request("VendedorID")
Zona = Trim(Request("Zona"))
ObservacionesAdicional = Trim(Request("ObservacionesAdicional"))


' ================================
' Orden entregada y fecha entregada
' ================================
If Trim(Request("OrdenEntregada")) = "" Then
    OrdenEntregada = 0
Else
    OrdenEntregada = CInt(Request("OrdenEntregada"))
	Estado = 3
End If

' ================================
' FechaEntrega (manejo de null)
' ================================
If Trim(Request("FechaEntrega")) = "" Then
    FechaEntrega = "NULL"
Else
    FechaEntrega = "'" & Request("FechaEntrega") & "'"
End If


If Trim(Request("FechaEntregada")) = "" Then
    FechaEntregada = "NULL"
Else
    FechaEntregada = "'" & Request("FechaEntregada") & "'"
End If


Dim sqlSP
sqlSP = ""
IF OrdenEntregada<>0 Then
' ================================
' Actualizar cabecera mediante SP
' ================================
	sqlSP = "EXEC report.[usp_Ordenes_Ventas_Administracion_FechaEntrega_Update] " & _
			FileID & ", " & _
			OrdenID & ", " & _ 
			Estado & ", " & _
			OrdenEntregada & ", " & _
			FechaEntregada

	'response.write sqlSP : response.end

	conn.Execute sqlSP
End If
 
' ================================
' Actualizar cabecera mediante SP
' ================================

sqlSP = ""
sqlSP = "EXEC report.usp_Ordenes_Ventas_Update " & _
        FileID & ", " & _
        OrdenID & ", " & _
        "'" & VendedorID & "', " & _
        FechaEntrega & ", " & _
        "'" & Replace(Zona,"'","''") & "', " & _
        "'" & Replace(ObservacionesAdicional,"'","''") & "' ," & _
        1 
		
conn.Execute sqlSP

' ================================
' Actualizar cajas modificadas mediante SP
' ================================
Dim sqlCajas, rsCajas
sqlCajas = "EXEC report.usp_Ordenes_Ventas_Cajas_sel " & FileID & "," & OrdenID
'response.write sqlCajas
Set rsCajas = conn.Execute(sqlCajas)

Do While Not rsCajas.EOF
    Dim CajaID, checkName, Modificado
    CajaID = rsCajas("CajaID")
    checkName = "CajaMod_" & CajaID

    If Request(checkName) <> "" Then
        Modificado = 1
    Else
        Modificado = 0
    End If
	'response.write "EXEC report.usp_Ordenes_Ventas_Cajas_Update_Modificado " & FileID & "," & OrdenID  & "," & CajaID & "," & Modificado
    ' Llamar SP para actualizar Modificado
    conn.Execute "EXEC report.usp_Ordenes_Ventas_Cajas_Update_Modificado " & CajaID & "," & Modificado

    rsCajas.MoveNext
Loop

rsCajas.Close
Set rsCajas = Nothing
' ================================
' Redirigir a ventas_detalle
' ================================
Response.Redirect "ventas_detalle.asp?FechaInicio=" & toISO(FechaInicio) & "&FechaFin= " & toISO(FechaFin) & "&order=" & orderBy & "&dir=" & orderDir &"&mensaje=Orden+guardada+correctamente"
%>
