<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

Function FechaLatina(fecha)
    If IsDate(fecha) Then
        FechaLatina = Day(fecha) & "/" & Right("0" & Month(fecha),2) & "/" & Year(fecha)
    Else
        FechaLatina = fecha
    End If
End Function 

Dim vehiculoID
vehiculoID = Trim(Request("vehiculoID"))
If vehiculoID = "" Then
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "dbo.usp_Vehiculos_Mantenimiento_Historial"
cmd.Parameters.Append cmd.CreateParameter("@VehiculoID", 3, 1, , CLng(vehiculoID))

Set rs = cmd.Execute()

If rs.EOF Then
    Response.Write "<div style='color:#6b7280;font-size:13px;'>No hay comentarios registrados para este vehículo.</div>"
Else
    Do While Not rs.EOF
%>
        <div class="timeline-item">
            <div class="timeline-header">
                <div class="timeline-date">
                    <%=FechaLatina(rs("Fecha"))%> - <strong><%=Server.HTMLEncode(rs("Usuario"))%></strong>
                </div>
                <div>
                    <button class="btn btn-blue" onclick="editar(<%=rs("ID")%>)">Editar</button>
                    <button class="btn btn-red" onclick="eliminarComentario(<%=rs("ID")%>)">Eliminar</button>
                </div>
            </div>
            <div class="timeline-text" id="txt_<%=rs("ID")%>">
                <%=Server.HTMLEncode(rs("Comentario"))%>
            </div>
        </div>
<%
        rs.MoveNext
    Loop
End If

rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>
