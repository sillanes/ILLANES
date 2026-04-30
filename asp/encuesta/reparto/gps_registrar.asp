<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
On Error Resume Next

Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.Status = "200 OK"

' ============================================================
' PARÁMETROS
' ============================================================
Dim transportistaid, lat, lon, fechagps,origen, accuracy, speed, provider, hdrid
transportistaid = Trim(Request("transportistaid"))
lat  = Trim(Request("lat"))
lon  = Trim(Request("lon"))
fechagps = Trim(Request("fechagps"))
accuracy = Request("accuracy")
speed = Request("speed")
provider = Request("provider")
hdrid = Request("hojaderutaid")

origen = Trim(Request("origen"))
If origen = "" Then origen = "ANDROID"
' ============================================================
' LIMPIAR ERRORES
' ============================================================
Err.Clear
conn.Errors.Clear


If accuracy = "" Then accuracy = 0
If speed = "" Then speed = 0
If provider = "" Then provider = "unknown"
If origen = "" Then origen = "ANDROID"

' ============================================================
' EJECUTAR STORED PROCEDURE (CONEXIÓN PRINCIPAL)
' ============================================================
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "tracking.usp_Transportista_GPS_Ins"

cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , transportistaid)
cmd.Parameters.Append cmd.CreateParameter("@HojaDeRutaID", 3, 1, , hdrid )
cmd.Parameters.Append cmd.CreateParameter("@Latitud", 200, 1, 20, lat)
cmd.Parameters.Append cmd.CreateParameter("@Longitud", 200, 1, 20, lon)
cmd.Parameters.Append cmd.CreateParameter("@PrecisionMetros", 3, 1, , 0)
cmd.Parameters.Append cmd.CreateParameter("@FechaHoraGPS", 200, 1, 19, fechagps)
cmd.Parameters.Append cmd.CreateParameter("@Origen", 200, 1, 20, origen)
cmd.Parameters.Append cmd.CreateParameter("@DispositivoID", 200, 1, 100, "ANDROID")
cmd.Parameters.Append cmd.CreateParameter("@Bateria", 3, 1, , 0)
cmd.Parameters.Append cmd.CreateParameter("@Velocidad", 200, 1, 10, "") 
cmd.Parameters.Append cmd.CreateParameter("@Accuracy", 200, 1, 10, accuracy)
cmd.Parameters.Append cmd.CreateParameter("@Speed", 200, 1, 10, speed)
cmd.Parameters.Append cmd.CreateParameter("@Provider", 200, 1, 20, provider)



cmd.Execute

' ============================================================
' SI HUBO ERROR SQL → LOGUEAR CON OTRA CONEXIÓN
' ============================================================
If conn.Errors.Count > 0 Then

    Dim logConn
    Set logConn = Server.CreateObject("ADODB.Connection")
    logConn.Open conn.ConnectionString   ' MISMO CONNECTION STRING

    Dim adoErr
    For Each adoErr In conn.Errors
        Dim sqlLog
        sqlLog = "INSERT INTO tracking.GPS_Errores_Log " & _
                 "(TransportistaID, Latitud, Longitud, FechaHoraGPS, ErrorNumero, ErrorDescripcion, Origen) VALUES (" & _
                 transportistaid & ", " & _
                 "'" & Replace(lat,"'","") & "', " & _
                 "'" & Replace(lon,"'","") & "', " & _
                 "'" & Replace(fechagps,"'","") & "', " & _
                 adoErr.Number & ", " & _
                 "'" & Replace(adoErr.Description,"'","") & "', " & _
                 "'SQL')"
        logConn.Execute sqlLog
    Next

    logConn.Close
    Set logConn = Nothing

    conn.Errors.Clear
End If

Response.Write "OK"
Response.End
%>
