<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

Dim hdrid
hdrid = Request("hdrid")

' --- Procesar envío del formulario ---
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim importe, gastos, obsDeposito, obsGastos
    importe     = CDbl(Request.Form("importe"))
    gastos      = CDbl(Request.Form("gastos"))
    obsDeposito = Replace(Request.Form("obsDeposito"),"'", "''")
    obsGastos   = Replace(Request.Form("obsGastos"),"'", "''")

    Dim sql
    sql = "EXEC usp_Transportista_HojaDeRuta_Cerrar " & hdrid & ", " & _
          Replace(importe,",",".") & ", " & _
          Replace(gastos,",",".") & ", '" & obsDeposito & "', '" & obsGastos & "'"
	'response.write sql
    On Error Resume Next
    set rs = conn.Execute (sql)
	
    If Err.Number = 0 Then
		If Not rs.EOF Then
			hasError = rs("hasError")
			ErrorMessage = rs("ErrorMessage")
		End If
		rs.Close
		Set rs=Nothing
		If hasError=0 Then
			Response.Redirect "hojaderutaV3.asp?msg=ok"
		End If 
		
    Else
        'Response.Write "<div style='color:red;padding:10px;'>Error al cerrar hoja de ruta: " & Err.Description & "</div>"
        Err.Clear
    End If
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cerrar Hoja de Ruta</title>
    <link rel="stylesheet" href="estilos.css">
	
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        .form-container {
            max-width: 600px;
            margin: 40px auto;
            padding: 20px;
            background: #f7f7f7;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }
        label { font-weight: bold; display: block; margin-bottom: 6px; }
        .row {
            display: flex;
            gap: 10px;
            margin-bottom: 16px;
        }
        .row > div { flex: 1; }
        input[type="text"] {
            width: 100%;
            padding: 8px;
            border-radius: 6px;
            border: 1px solid #ccc;
        }
        .btn {
            background: #27ae60;
            color: white;
            border: none;
            padding: 10px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
        }
        .btn:hover { background: #1f8a4c; }
        .btn-back {
            background: #2e30c7;
            margin-left: 8px;
            text-decoration: none;
            display: inline-block;
            padding: 10px 16px;
            border-radius: 6px;
            color: white;
        }
    </style>
	<script>

	function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}
	</script>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex:1;">👤 <%=Server.HTMLEncode(Session("NombreTransportista"))%></strong>
    <form method="post" action="logout.asp" style="margin:0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content">

	<% If hasError = 1 Then %>
	
<div class="container mt-5">
	
<div class="container mt-5">
  <div class="alert alert-danger p-4 rounded shadow">
    <h4 class="alert-heading">Error al cerrar hoja de ruta </h4>
    <p><%= ErrorMessage %></p>
    <hr> 
  </div>
</div>
</div>


	<% End If %>
<div class="form-container">
    <h2>Cerrar Hoja de Ruta <%=hdrid%></h2>


    <form method="post" action="hojaderutav3_cerrar.asp?hdrid=<%=hdrid%>">
        <div class="row">
            <div>
                <label for="importe">Importe Depositado:</label>
                <input type="text" id="importe" name="importe" value="0.00" required>
            </div>
            <div>
                <label for="obsDeposito">Observaciones Depósito:</label>
                <input type="text" id="obsDeposito" name="obsDeposito" maxlength="255">
            </div>
        </div>

        <div class="row">
            <div>
                <label for="gastos">Gastos:</label>
                <input type="text" id="gastos" name="gastos" value="0.00" required>
            </div>
            <div>
                <label for="obsGastos">Observaciones Gastos:</label>
                <input type="text" id="obsGastos" name="obsGastos" maxlength="255">
            </div>
        </div>

        <div style="display:flex;justify-content:flex-end;">
            <button type="submit" class="btn">Guardar</button>
            <a href="hojaderutaV3.asp?hdrid=<%=hdrid%>" class="btn-back">Volver</a>
        </div>
    </form>
</div>
</div>
</body>
</html>
