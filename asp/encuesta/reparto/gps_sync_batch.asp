<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim json, transportistaID, dispositivoID
json = Request.Form("json")

If json = "" Then
    Response.Write "ERROR: JSON VACIO"
    Response.End
End If

' =====================================================
' PARSEO JSON (VBScript manual o lib JSON)
' =====================================================
' Recomendado: https://github.com/VBA-tools/VBA-JSON
' Se asume que ya lo usás o lo incorporás

Dim data, ubicacion, cmd
Set data = ParseJson(json)

transportistaID = data("transportistaid")
dispositivoID   = data("dispositivoid")

For Each ubicacion In data("ubicaciones")

    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = 4
    cmd.CommandText = "tracking.usp_Transportista_GPS_Ins"

    cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , transportistaID)
    cmd.Parameters.Append cmd.CreateParameter("@HojaDeRutaID", 3, 1, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@Latitud", 5, 1, , Replace(ubicacion("lat"),",","."))
    cmd.Parameters.Append cmd.CreateParameter("@Longitud", 5, 1, , Replace(ubicacion("lon"),",","."))
    cmd.Parameters.Append cmd.CreateParameter("@PrecisionMetros", 3, 1, , ubicacion("precision"))
    cmd.Parameters.Append cmd.CreateParameter("@FechaHoraGPS", 7, 1, , ubicacion("fechagps"))
    cmd.Parameters.Append cmd.CreateParameter("@Origen", 200, 1, 20, "OFFLINE_SYNC")
    cmd.Parameters.Append cmd.CreateParameter("@DispositivoID", 200, 1, 100, dispositivoID)
    cmd.Parameters.Append cmd.CreateParameter("@Bateria", 3, 1, , ubicacion("bateria"))
    cmd.Parameters.Append cmd.CreateParameter("@Velocidad", 5, 1, , ubicacion("velocidad"))

    cmd.Execute
Next

Response.Write "OK"
%>
