<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="_upload.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

  
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function
If Session("NombreTransportista") = "" Then
  Response.Redirect "login.asp"
End If

Dim hdrid, clienteid, msg
hdrid     = Request("hdrid")
clienteid = Request("clienteid")
dowhat = clng("0"&Request("dowhat"))
msg = ""

Dim fromPage, returnUrl
fromPage = LCase(Trim(Request("from")))
returnUrl = Trim(Request("return"))


If hdrid = "" Or clienteid = "" Then
  Response.Write "Datos incompletos"
  Response.End
End If


 Function do_Upload(rel_Folder)
  Server.ScriptTimeout = 3600
  Dim Form: Set Form = New ASPForm 

  'Form.UploadID = Request.QueryString("UploadID")
  Form.UploadID = Request("UploadID")
  Form.SizeLimit = 10*1024*1024 '10MB

  Dim HTML, hResult
  Const fsCompletted  = 0
  Const fsSizeLimit   = &HD
  Const fsTimeOut     = &HE
  Const fsError       = &HA
  
  If Form.State > fsError Then 
    If Form.State = fsSizeLimit Then 
      hResult = "El archivo supera el límite permitido (" & Form.SizeLimit/1024 & " KB)."
    ElseIf Form.State = fsTimeOut Then
      hResult = "El tiempo de carga excedió el máximo permitido (" & Form.ReadTimeout & " seg)."
    Else
      hResult = "Error de carga (código " & Form.State & ")."
    End If
    hResult = "<div class='textError'>" & hResult & "</div>"
    Response.Status = "400 Bad request"
ElseIf Form.State = fsCompletted Then

  Session("FileUploaded") = nombreFinal 
End If


  Dim UploadID
  UploadID = Form.NewUploadID

  HTML = ""
  HTML = HTML & hResult
  HTML = HTML & "<form name='FFUP' method='post' enctype='multipart/form-data' " & _
                 "action='" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&Step=2&clienteid=" & clienteid & "&hdrid=" & hdrid) & "' " & _
                 "onsubmit='return ProgressBar();'>"
  HTML = HTML & "<div class='form-group'>"
  HTML = HTML & "  <label for='File1'>Seleccione un archivo:</label>"
  'HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required  accept='.pdf,.jpg,.jpeg,.png'  capture='environment'>"
  HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required  accept='.pdf,.jpg,.jpeg,.png'>"
  HTML = HTML & "</div>"
  HTML = HTML & "<div class='form-actions'>"
  HTML = HTML & "  <button type='submit' class='btn-volver'>Subir Archivo</button>"
  
  HTML = HTML & "</div>"
  HTML = HTML & "</form>"

  HTML = HTML & "<script>" & vbCrLf
  HTML = HTML & "function ProgressBar(){" & vbCrLf
  HTML = HTML & "  var ProgressURL = 'progress.asp?UploadID=" & UploadID & "';" & vbCrLf
  HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
  HTML = HTML & "  return true;" & vbCrLf
  HTML = HTML & "}" & vbCrLf
  HTML = HTML & "</script>"

  do_Upload = HTML
End Function

Function do_Upload_PDF(rel_Folder)
  Server.ScriptTimeout = 3600

  Dim Form: Set Form = New ASPForm

  ' ✅ NUNCA uses Request("UploadID") en multipart
  Dim upid
  upid = Trim(Request.QueryString("UploadID"))
  If upid <> "" Then Form.UploadID = upid

  Form.SizeLimit = 10*1024*1024 '10MB (si fotos pesan más, subilo a 20/30MB)

  ' ✅ para móviles (si el componente soporta ReadTimeout)
  On Error Resume Next
  Form.ReadTimeout = 600
  On Error GoTo 0

  Dim HTML, hResult
  Const fsCompletted  = 0
  Const fsSizeLimit   = &HD
  Const fsTimeOut     = &HE
  Const fsError       = &HA

  If Form.State > fsError Then
    If Form.State = fsSizeLimit Then
      hResult = "El archivo supera el límite permitido (" & Form.SizeLimit/1024 & " KB)."
    ElseIf Form.State = fsTimeOut Then
      hResult = "El tiempo de carga excedió el máximo permitido (" & Form.ReadTimeout & " seg)."
    Else
      hResult = "Error de carga (código " & Form.State & ")."
    End If
    hResult = "<div class='textError'>" & hResult & "</div>"
    Response.Status = "400 Bad request"

  ElseIf Form.State = fsCompletted Then
    Form.Files.Save MapFolderToDisk(rel_Folder)
    hResult = "<div class='textSuccess'>Archivo subido correctamente.</div>"

  ElseIf Request.QueryString("Action") = "Cancel" Then
    hResult = "<div class='textWarning'>La carga fue cancelada.</div>"
  End If

  Dim UploadID
  UploadID = Form.NewUploadID

  HTML = ""
  HTML = HTML & hResult
  HTML = HTML & "<form name='FFUP' method='post' enctype='multipart/form-data' " & _
                "action='" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&Step=2&clienteid=" & clienteid & "&hdrid=" & hdrid) & "' " & _
                "onsubmit='return ProgressBarSmart();'>"

  HTML = HTML & "<div class='form-group'>"
  HTML = HTML & "  <label for='File1'>Seleccione un archivo:</label>"
  ' ✅ mejor para móviles: permitir formatos modernos también
  HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required accept='image/*,application/pdf'>"
  HTML = HTML & "</div>"

  HTML = HTML & "<div class='form-actions'>"
  HTML = HTML & "  <button type='submit' class='btn-volver'>Subir Archivo</button>"
  HTML = HTML & "</div>"
  HTML = HTML & "</form>"

  HTML = HTML & "<script>" & vbCrLf
  HTML = HTML & "function isMobile(){ return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent); }" & vbCrLf
  HTML = HTML & "function ProgressBarSmart(){ if(isMobile()){ return true; }" & vbCrLf
  HTML = HTML & "  var ProgressURL='progress.asp?UploadID=" & UploadID & "';" & vbCrLf
  HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
  HTML = HTML & "  return true; }" & vbCrLf
  HTML = HTML & "</script>"

  do_Upload_PDF = HTML
End Function


step=0


if doWhat="" or doWhat<"2"   Then
	 Session("FileUploaded") = ""
End If

%>

<!-- ===============================
     HTML
     =============================== -->
<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Comprobantes</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css">


<style>
.container{max-width:900px;margin:20px auto}
.card{border:1px solid #ddd;border-radius:6px;padding:15px;margin-bottom:20px;background:#fff}
.table{width:100%;border-collapse:collapse}
.table th,.table td{border:1px solid #ddd;padding:8px;text-align:left}
.table th{background:#f2f2f2}
.alert{padding:10px;margin-bottom:15px;border-radius:4px}
.alert-ok{background:#d4edda;color:#155724}
.alert-error{background:#f8d7da;color:#721c24}
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

<div class="container">
 

<div class="d-flex justify-content-between align-items-center mb-4">
  <h2>Hoja de Ruta <%=hdrid%> – Cliente <%=clienteid%></h2> 
  <a href="hojaderutav3.asp?hdrid=<%= hdrid %>"  
     class="icon-btn"
     style="background:#2e30c7;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none">
    <i class="fas fa-arrow-left"></i> Volver
  </a>
</div>


<% If msg <> "" Then %>
<div class="alert <%=IIf(InStr(msg,"correctamente")>0,"alert-ok","alert-error")%>">
  <%=msg%>
</div>
<% End If %>



<%
' PASO 4: Subir archivo PDF
If step = "0" Then 
    msg = ""
%>

<h2>Adjuntar comprobante</h2>
<% 
			
	If doWhat >= 0 Then
 
		if doWhat = 3 then
			%>
			<div style="display:none">
			<%
		Else
			%>
			<div style="display:block">
			<%
		End If
	%>
		 
 		<%=do_Upload_PDF("")%>
		<!--#INCLUDE FILE="_commonXLS.asp"--> 
		</div>
	<%
	End if
	 
		if Session("FileUploaded") <> "" then

			%><!--#include file="./includes/hojaderuta_comprobantes_ins.asp" --><%
			

			
	%> 
	
		
			<div style="text-align:center; margin:20px 0;">
			<span class="textSuccess" 
			      style="display:inline-block; padding:12px 20px; 
			             background:#28a745; color:white; 
			             border-radius:25px; font-weight:bold;">
				📂 Archivo subido correctamente:<br>
				<%=Session("FileUploaded")%> 
				
				<%Session("FileUploaded") = ""%>
			</span>
		</div>
		 
		
		<div class="form-actions" style="text-align:center; margin-top:25px;">
		
		<button type="button"
		  onclick="window.location.href='hojaderuta_comprobantes.asp?hdrid=<%=hdrid%>&clienteid=<%=clienteid%>&step=1'"
		  class="btn btn-outline-primary"
		  style="font-size:16px; padding:10px 18px;">
		  <i class="fas fa-plus-circle"></i> Agregar otro comprobante
		</button>

			
        </div>
 
		   

		 	
		
		 
	<%
		Else

		End If 
	%>

<%
End If
%>

<div class="card">
<h3>📂 Comprobantes cargados</h3>

<table class="table">
<tr>
  <th>Fecha</th>
  <th>Archivo</th>
  <th>Tipo</th>
  <th>Acción</th>
</tr>

<%
Dim rs
Set rs = conn.Execute("EXEC cobranza.Transportista_Comprobante_Sel " & hdrid & "," & clienteid)

If rs.EOF Then
%>
<tr><td colspan="4">No hay comprobantes cargados</td></tr>
<%
Else
Do Until rs.EOF
%>
<tr>
  <td><%=rs("FechaSubida")%></td>
  <td><%=rs("NombreArchivo")%></td>
  <td><%=rs("Extension")%></td>

  
 <td style="white-space:nowrap;">
  <a href="<%= "https://illanes-encuesta.ddns.net/comprobantes/" & rs("NombreArchivo")%>"
     target="_blank"
     title="Ver comprobante">
    <i class="fas fa-eye"></i>
  </a><a href="hojaderuta_comprobantes_del.asp?id=<%=rs("ComprobanteID")%>&hdrid=<%=hdrid%>&clienteid=<%=clienteid%>"
     onclick="return confirm('¿Eliminar este comprobante?');"
     title="Eliminar comprobante"
     style="color:#dc3545;">
    <i class="fas fa-trash-alt"></i>
  </a>
</td>

 
</tr>
<%
rs.MoveNext
Loop
End If

rs.Close
Set rs = Nothing
%>

</table>
</div>

</div>

<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
