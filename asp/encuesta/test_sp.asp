<%
' Script de prueba para el stored procedure usp_Clientes_Alta
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim dbCon, cmd, altaClienteID
Dim testResult

testResult = "INICIANDO PRUEBA...<br>"

On Error Resume Next

Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_command_const.asp" --><%
%><!--#include file="./includes/db_con_open_ventas.asp" --><%

If Err.Number <> 0 Then
    testResult = testResult & "ERROR DE CONEXIÓN: " & Err.Description & "<br>"
    Err.Clear
Else
    testResult = testResult & "CONEXIÓN EXITOSA<br>"

    ' Verificar si existe la tabla
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT COUNT(*) as Count FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Clientes_Alta'", dbCon
    If rs("Count") > 0 Then
        testResult = testResult & "TABLA Clientes_Alta EXISTE<br>"
    Else
        testResult = testResult & "TABLA Clientes_Alta NO EXISTE<br>"
    End If
    rs.Close
    Set rs = Nothing

    ' Verificar si existe el SP
    rs.Open "SELECT COUNT(*) as Count FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_NAME = 'usp_Clientes_Alta'", dbCon
    If rs("Count") > 0 Then
        testResult = testResult & "STORED PROCEDURE usp_Clientes_Alta EXISTE<br>"
    Else
        testResult = testResult & "STORED PROCEDURE usp_Clientes_Alta NO EXISTE<br>"
    End If
    rs.Close
    Set rs = Nothing

    ' Intentar ejecutar el SP con datos de prueba
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = dbCon
    cmd.CommandType = adCmdStoredProcedure
    cmd.CommandText = "dbo.usp_Clientes_Alta"

    cmd.Parameters.Append cmd.CreateParameter("@VendedorID", adParamVarchar, adParamInput, 20, "TEST001")
    cmd.Parameters.Append cmd.CreateParameter("@VendedorNombre", adParamVarchar, adParamInput, 120, "Vendedor de Prueba")
    cmd.Parameters.Append cmd.CreateParameter("@Nombre", adParamVarchar, adParamInput, 80, "Juan")
    cmd.Parameters.Append cmd.CreateParameter("@Apellido", adParamVarchar, adParamInput, 80, "Pérez")
    cmd.Parameters.Append cmd.CreateParameter("@Direccion", adParamVarchar, adParamInput, 160, "Calle Falsa 123")
    cmd.Parameters.Append cmd.CreateParameter("@CUITCUIL", adParamVarchar, adParamInput, 20, "20123456789")
    cmd.Parameters.Append cmd.CreateParameter("@Provincia", adParamVarchar, adParamInput, 80, "Buenos Aires")
    cmd.Parameters.Append cmd.CreateParameter("@Ciudad", adParamVarchar, adParamInput, 80, "La Plata")
    cmd.Parameters.Append cmd.CreateParameter("@ConstanciaAfipArchivo", adParamVarchar, adParamInput, 255, "test.pdf")
    cmd.Parameters.Append cmd.CreateParameter("@ConstanciaAfipRuta", adParamVarchar, adParamInput, 500, "C:\test\test.pdf")
    cmd.Parameters.Append cmd.CreateParameter("@AltaClienteID", adParamInt, adParamOutput)

    cmd.Execute

    If Err.Number <> 0 Then
        testResult = testResult & "ERROR AL EJECUTAR SP: " & Err.Description & "<br>"
        Err.Clear
    Else
        altaClienteID = cmd.Parameters("@AltaClienteID").Value
        testResult = testResult & "STORED PROCEDURE EJECUTADO EXITOSAMENTE<br>"
        testResult = testResult & "ID GENERADO: " & altaClienteID & "<br>"
    End If

    Set cmd = Nothing
End If

If Not (dbCon Is Nothing) Then
    dbCon.Close
    Set dbCon = Nothing
End If

On Error GoTo 0

testResult = testResult & "PRUEBA FINALIZADA"
%>

<!DOCTYPE html>
<html>
<head>
    <title>Prueba Stored Procedure</title>
    <meta charset="UTF-8">
</head>
<body>
    <h1>Resultado de Prueba del Stored Procedure</h1>
    <div style="font-family: monospace; background: #f5f5f5; padding: 20px; border: 1px solid #ccc;">
        <%=testResult%>
    </div>
</body>
</html>