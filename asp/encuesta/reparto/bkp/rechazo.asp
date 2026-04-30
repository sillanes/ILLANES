<!--#include file="header.asp" -->

<%
Dim hojaRutaID, motivo, msg
hojaRutaID = Request.QueryString("ID")
motivo = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    motivo = Request.Form("motivo")
    If motivo = "" Then
        msg = "<div class='alert alert-danger'>Debe ingresar un motivo de rechazo.</div>"
    Else
        ' Insertar motivo en base de datos, ejemplo:
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = "EXEC usp_InsertarMotivoRechazo ?, ?"
        cmd.Parameters.Append cmd.CreateParameter("@HojaRutaID", 3, 1, , hojaRutaID)
        cmd.Parameters.Append cmd.CreateParameter("@Motivo", 200, 1, 500, motivo)
        cmd.Execute
        Set cmd = Nothing

        Response.Redirect "home.asp"
    End If
End If
%>

<h1>Registrar Motivo de Rechazo</h1>

<%= msg %>

<form method="post" action="rechazo.asp?ID=<%= hojaRutaID %>">
    <div class="mb-3">
        <label for="motivo" class="form-label">Motivo</label>
        <textarea class="form-control" id="motivo" name="motivo" rows="4"><%= motivo %></textarea>
    </div>
    <button type="submit" class="btn btn-danger">Guardar motivo y volver</button>
    <a href="home.asp" class="btn btn-secondary ms-2">Cancelar</a>
</form>

<!--#include file="footer.asp" -->
