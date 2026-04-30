<%
sSQL = "EXEC report.usp_Pxp_Eficiencia_Mes 0,1"

Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.Open sSQL, dbCon

If Err.Number <> 0 Then
    Response.Clear
    Response.Redirect "./error.asp"
End If

Dim esAdminRRHH
esAdminRRHH = False

If Session("currentUser") = "admin" Or Session("currentUser") = "rrhh" Then
    esAdminRRHH = True
End If
%>

<div class="rep-list">

<%
If Not dbRS.EOF Then

    Do Until dbRS.EOF
%>

    <div class="rep-row-card">

        <div class="rep-row-main">
            <div class="rep-periodo">
                <span class="rep-label">Año</span>
                <strong><%=dbRS("PeriodoAnio")%></strong>
            </div>

            <div class="rep-periodo">
                <span class="rep-label">Periodo</span>
                <strong><%=dbRS("PeriodoNombre")%></strong>
            </div>
        </div>

        <div class="rep-metrics">

            <div class="rep-metric">
                <span class="rep-label">Pedidos</span>
                <strong><%=dbRS("CantPedidos")%></strong>
            </div>

            <div class="rep-metric">
                <span class="rep-label">Errores</span>
                <strong><%=dbRS("CantErrores")%></strong>
            </div>

            <div class="rep-metric">
                <span class="rep-label">Eficiencia</span>
                <strong><%=FormatNumber(dbRS("Eficiencia"),2)%>%</strong>
            </div>

            <% If esAdminRRHH Then %>

            <div class="rep-metric money">
                <span class="rep-label">Premio</span>
                <strong>$ <%=dbRS("Premio")%></strong>
            </div>

            <div class="rep-metric money total">
                <span class="rep-label">Total</span>
                <strong>$ <%=FormatNumber(dbRS("Total"),2)%></strong>
            </div>

            <% End If %>

        </div>

    </div>

<%
        dbRS.MoveNext
    Loop

Else
%>

    <div class="rep-empty">
        No hay datos
    </div>

<%
End If
%>

</div>

<%
Set dbRS = Nothing
%>