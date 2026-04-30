<%@ Language="VBScript" %>
<!--#include file="../empretienda/conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
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

' ================================
' Variables recibidas por POST
' ================================
Dim OrdenID, FileID, VendedorID, FechaEntrega, ObservacionesAdicional
Dim Envio_Calle, Envio_Numero, Envio_Piso, Envio_Dpto, Envio_Ciudad
Dim Estado, OrdenEntregada, FechaEntregada

FileID = Request("FileID")
OrdenID = Request("OrdenID")
VendedorID = Request("VendedorID") 
ObservacionesLogistica = Trim(Request("ObservacionesLogistica"))

Envio_Calle  = Replace(Trim(Request("Envio_Calle")),"'","''")
Envio_Numero = Replace(Trim(Request("Envio_Numero")),"'","''")
Envio_Piso   = Replace(Trim(Request("Envio_Piso")),"'","''")
Envio_Dpto   = Replace(Trim(Request("Envio_Dpto")),"'","''")
Envio_Ciudad = Replace(Trim(Request("Envio_Ciudad")),"'","''")

' ================================
' FechaEntrega (manejo de null)
' ================================
If Trim(Request("FechaEntrega")) = "" Then
    FechaEntrega = "NULL"
	Estado = 0
Else
    FechaEntrega = "'" & Request("FechaEntrega") & "'"
	Estado = 2
End If

' ================================
' Orden entregada y fecha entregada
' ================================
If Trim(Request("OrdenEntregada")) = "" Then
    OrdenEntregada = 0
Else
    OrdenEntregada = CInt(Request("OrdenEntregada"))
	Estado = 3
End If

If Trim(Request("FechaEntregada")) = "" Then
    FechaEntregada = "NULL"
Else
    FechaEntregada = "'" & Request("FechaEntregada") & "'"
End If

' ================================
' Actualizar cabecera mediante SP
' ================================
Dim sqlSP
sqlSP = "EXEC report.[usp_Ordenes_Ventas_Logistica_Update] " & _
        FileID & ", " & _
        OrdenID & ", " & _
        "'" & VendedorID & "', " & _
        FechaEntrega & ", " & _
        "'" & ObservacionesLogistica & "', " & _
        "'" & Envio_Calle & "', " & _
        "'" & Envio_Numero & "', " & _
        "'" & Envio_Piso & "', " & _
        "'" & Envio_Dpto & "', " & _
        "'" & Envio_Ciudad & "', " & _
        Estado & ", " & _
        OrdenEntregada & ", " & _
        FechaEntregada

'response.write sqlSP : response.end

conn.Execute sqlSP

' ================================
' Actualizar cajas modificadas mediante SP
' ================================
'Dim sqlCajas, rsCajas
'sqlCajas = "EXEC report.usp_Ordenes_Ventas_Cajas_sel " & FileID & "," & OrdenID
'Set rsCajas = conn.Execute(sqlCajas)
'
'Do While Not rsCajas.EOF
'    Dim CajaID, checkName, Modificado
'    CajaID = rsCajas("CajaID")
'    checkName = "CajaMod_" & CajaID
'
'    If Request(checkName) <> "" Then
'        Modificado = 1
'    Else
'        Modificado = 0
'    End If
'
'    conn.Execute "EXEC report.usp_Ordenes_Ventas_Cajas_Update_Modificado " & CajaID & "," & Modificado
'    rsCajas.MoveNext
'Loop
'
'rsCajas.Close
'Set rsCajas = Nothing

' ================================
' Redirigir a ventas_detalle
' ================================
'Response.Redirect "ordenes_detalle.asp?FechaInicio=" & toISO(FechaInicio) & "&FechaFin=" & toISO(FechaFin) & "&mensaje=Orden+guardada+correctamente"
Response.Redirect "ordenes_modificar.asp?FileID=" & FileID &"&OrdenID="& OrdenID & "&FechaInicio=" & toISO(FechaInicio) & "&FechaFin=" & toISO(FechaFin) & "&mensaje=Orden+guardada+correctamente"
%>
