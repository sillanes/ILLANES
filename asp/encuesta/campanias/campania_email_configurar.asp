<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->

<!--#INCLUDE FILE="_upload.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then Response.Redirect "../login.asp"


Dim CampaniaID, step, doWhat
CampaniaID = Request.QueryString("CampaniaID")
doWhat = Request.QueryString("doWhat") 

If CampaniaID = "" Then CampaniaID = 1
step = Request.QueryString("step")
If step = "" Then step = "1"


  
 Function do_Upload(rel_Folder)
  Server.ScriptTimeout = 3600
  Dim Form: Set Form = New ASPForm 

  Form.UploadID = Request.QueryString("UploadID")
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
                 "action='" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&Step=2&CampaniaID=" & CampaniaID) & "' " & _
                 "onsubmit='return ProgressBar();'>"
  HTML = HTML & "<div class='form-group'>"
  HTML = HTML & "  <label for='File1'>Seleccione un archivo Excel:</label>"
  HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required>"
  HTML = HTML & "</div>"
  HTML = HTML & "<div class='form-actions'>"
  HTML = HTML & "  <button type='submit' class='btn-volver'>Subir Archivo</button>"
  HTML = HTML & "  <button type='button' onclick='window.location.href=""campania_email_configurar.asp?CampaniaID=" & CampaniaID & """' class='btn-cancelar'>Cancelar</button>"
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

  Form.UploadID = Request.QueryString("UploadID")
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
                 "action='" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3&Step=4&CampaniaID=" & CampaniaID) & "' " & _
                 "onsubmit='return ProgressBar();'>"
  HTML = HTML & "<div class='form-group'>"
  HTML = HTML & "  <label for='File1'>Seleccione un archivo PDF:</label>"
  HTML = HTML & "  <input type='file' id='File1' name='File1' class='input-file' required>"
  HTML = HTML & "</div>"
  HTML = HTML & "<div class='form-actions'>"
  HTML = HTML & "  <button type='submit' class='btn-volver'>Subir Archivo</button>"
  HTML = HTML & "  <button type='button' onclick='window.location.href=""campania_email_configurar.asp?CampaniaID=" & CampaniaID & """' class='btn-cancelar'>Cancelar</button>"
  HTML = HTML & "</div>"
  HTML = HTML & "</form>"

  HTML = HTML & "<script>" & vbCrLf
  HTML = HTML & "function ProgressBar(){" & vbCrLf
  HTML = HTML & "  var ProgressURL = 'progress.asp?UploadID=" & UploadID & "';" & vbCrLf
  HTML = HTML & "  window.open(ProgressURL,'_blank','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200');" & vbCrLf
  HTML = HTML & "  return true;" & vbCrLf
  HTML = HTML & "}" & vbCrLf
  HTML = HTML & "</script>"

  do_Upload_PDF = HTML
End Function

  
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

Set objRequest = Nothing


if doWhat="" or doWhat<"3"   Then
	 Session("FileUploaded") = ""
End If

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Configurar Campaña</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
body { font-family: Arial, sans-serif; background:#f4f4f4; margin:0; padding:0; }
.main-content { max-width:900px; margin:30px auto; padding:20px; background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.1); }
h2 { text-align:center; margin-bottom:20px; }
.form-group { margin-bottom:15px; }
.form-group label { display:block; font-weight:bold; margin-bottom:5px; }
.form-group input, .form-group textarea { width:100%; padding:8px 10px; border:1px solid #ccc; border-radius:8px; font-size:14px; }
.form-actions { margin-top:20px; text-align:right; }
.btn-volver { padding:10px 20px; background:linear-gradient(135deg,#007bff,#0056b3); color:#fff; border:none; border-radius:25px; cursor:pointer; font-weight:bold; }
.btn-volver:hover { background:linear-gradient(135deg,#0056b3,#003f88); }
.table-container { overflow-x:auto; margin-top:20px; }
table { width:100%; border-collapse:collapse; }
th, td { border:1px solid #ddd; padding:8px; text-align:center; }
th { background:#007bff; color:#fff; }


.input-file {
  padding: 8px;
  border: 1px solid #ccc;
  border-radius: 12px;
  background: #fafafa;
  font-size: 14px;
  cursor: pointer;
}

.btn-cancelar {
  padding: 10px 20px;
  background: linear-gradient(135deg,#dc3545,#a71d2a);
  color: #fff;
  border: none;
  border-radius: 25px;
  cursor: pointer;
  font-weight: bold;
  margin-left: 10px;
}
.btn-cancelar:hover {
  background: linear-gradient(135deg,#a71d2a,#721c24);
}

.textError { color:#dc3545; font-weight:bold; margin:10px 0; }
.textSuccess { color:#28a745; font-weight:bold; margin:10px 0; }
.textWarning { color:#ffc107; font-weight:bold; margin:10px 0; }


</style>
</head>
<body>

<!--#include file="header.asp" -->
<div class="main-content">

	<form name="FF" method="post" action="campania_email_configuar.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	
 	
	</Form>
	
<%
' PASO 1: Selección / modificación de campaña
If step = "1" Then
    If Request.ServerVariables("REQUEST_METHOD")="POST" And Request.Form("action")="guardarPaso1" Then
        Dim nombreCampania, descripcion, sqlUpd
        nombreCampania = Replace(Request.Form("CampaniaNombre"),"'","''")
        descripcion = Replace(Request.Form("CampaniaDescripcion"),"'","''")
        sqlUpd = "EXEC dbo.usp_Campania_Step_0_upd " & CampaniaID & ",'" & nombreCampania & "','" & descripcion & "'"
        conn.Execute sqlUpd
        Response.Redirect "campania_email_configurar.asp?CampaniaID=" & CampaniaID & "&step=2"
    End If

    Dim rsDetalle, sql
    sql = "EXEC dbo.usp_Campania_Pendientes_Detalle_Sel " & CampaniaID
    Set rsDetalle = conn.Execute(sql)
%>
<h2>Paso 1: Información de la Campaña</h2>
<form method="post" action="campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=1">
    <div class="form-group">
        <label>Nombre de la campaña:</label>
        <input type="text" name="CampaniaNombre" value="<%=Server.HTMLEncode(rsDetalle("NombreCampania"))%>" required>
    </div>
    <div class="form-group">
        <label>Descripción:</label>
        <textarea name="CampaniaDescripcion" rows="4"><%=Server.HTMLEncode(rsDetalle("Descripcion"))%></textarea>
    </div>
    <div class="form-actions" style="text-align:center; margin-top:25px;">
        <button type="submit" name="action" value="guardarPaso1" class="btn-volver">Siguiente</button>
		<button type="button" onclick="window.location.href='campania_modificar.asp'" class="btn-cancelar">
			Volver
		</button>		
    </div>
</form>
<%
    rsDetalle.Close
    Set rsDetalle = Nothing
End If
%>

<%
' PASO 2: Subir archivo Excel
If step = "2" Then
    Dim upload, f, savePath, msg
    msg = ""
%>

<h2>Paso 2: Adjuntar excel con Emails</h2>
<% 
			
	if doWhat>="" then 
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
		 
 		<%=do_Upload("")%>
		<!--#INCLUDE FILE="_commonXLS.asp"--> 
		</div>
	<%
	End if
	
	
	If doWhat="3" Then

	
		if Session("FileUploaded") <> "" then

			
	%> 
	
			<div style="text-align:center; margin:20px 0;">
			<span class="textSuccess" 
			      style="display:inline-block; padding:12px 20px; 
			             background:#28a745; color:white; 
			             border-radius:25px; font-weight:bold;">
				📂 Archivo subido correctamente:<br>
				<%=Session("FileUploaded")%> 
			</span>
		</div>
		
		<!--#include file="./includes/TMP_campaniadel.asp" -->
		<!--#include file="./includes/TMP_campaniaxls.asp" -->
		<!--#include file="./includes/TMP_campaniains.asp" -->
		
		<div class="form-actions" style="text-align:center; margin-top:25px;">
            <button type="button" 
                    onclick="window.location.href='campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=3'" 
                    class="btn-volver">
                Siguiente
            </button>

			<button type="button" onclick="window.location.href='campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=1'" class="btn-cancelar">
				Volver
			</button>
			
        </div>
		 	
		
		 
	<%
		Else
	%>
		<div style="text-align:center; margin:20px;">
			<span style="color:#dc3545; font-weight:bold;">
				⚠️ No se recibió el archivo. Intenta de nuevo.
			</span>
		</div>
	<%
		End If
	End If
	%>

<%
End If
%>

<%
' PASO 3: Redacción de email
If step="3" Then
    If Request.ServerVariables("REQUEST_METHOD")="POST" And Request.Form("action")="guardarPaso3" Then
        Dim cuerpo, asunto
        cuerpo = Replace(Request.Form("CuerpoEmail"),"'","''")
        asunto = Replace(Request.Form("AsuntoEmail"),"'","''")
		sql = "EXEC usp_Campania_Step_3_upd " & CampaniaID & ", N'" & asunto & "', N'" & cuerpo & "'"
		conn.Execute sql 
		Response.Redirect "campania_email_configurar.asp?CampaniaID=" & CampaniaID & "&step=4"
		
    End If
%>
<h2>Paso 3: Redacción del Email</h2>
<form method="post">
    <div class="form-group">
        <label>Asunto:</label>
        <textarea name="AsuntoEmail" rows="1" required></textarea>
    </div>
	
    <div class="form-group">
        <label>Mensaje:</label>
        <textarea name="CuerpoEmail" rows="8" required></textarea>
    </div>	
 
    <div class="form-actions" style="text-align:center; margin-top:25px;">
        <button type="submit" name="action" value="guardarPaso3" class="btn-volver">Siguiente</button>
		<button type="button" onclick="window.location.href='campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=2'" class="btn-cancelar">
			Volver
		</button>		
    </div>
</form>
<%
End If
%>


<%
' PASO 4: Subir archivo PDF
If step = "4" Then 
    msg = ""
%>

<h2>Paso 4: Adjuntar archivo PDF</h2>
<% 
			
	if doWhat>="" then 
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
	
	
	If doWhat="3" Then

	
		if Session("FileUploaded") <> "" then

			
	%> 
	
		<!--#include file="./includes/campanias_pdf_update.asp" -->
		
			<div style="text-align:center; margin:20px 0;">
			<span class="textSuccess" 
			      style="display:inline-block; padding:12px 20px; 
			             background:#28a745; color:white; 
			             border-radius:25px; font-weight:bold;">
				📂 Archivo subido correctamente:<br>
				<%=Session("FileUploaded")%> 
			</span>
		</div>
		 
		
		<div class="form-actions" style="text-align:center; margin-top:25px;">
            <button type="button" 
                    onclick="window.location.href='campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=5'" 
                    class="btn-volver">
                Siguiente
            </button>

			<button type="button" onclick="window.location.href='campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=4'" class="btn-cancelar">
				Volver
			</button>
			
        </div>
		 	
		
		 
	<%
		Else
	%>
		<div style="text-align:center; margin:20px;">
			<span style="color:#dc3545; font-weight:bold;">
				⚠️ No se recibió el archivo. Intenta de nuevo.
			</span>
		</div>
	<%
		End If
	End If
	%>

<%
End If
%>


<%
' PASO 5: Cantidad por dia
If step="5" Then
    If Request.ServerVariables("REQUEST_METHOD")="POST" And Request.Form("action")="guardarPaso5" Then
        Dim batch 
        batch = Replace(Request.Form("batchsize"),"'","''") 
		sql = "usp_Campania_Step_5_upd " & CampaniaID & ",  "  & batch  
        conn.Execute(sql) 
		
		Response.Redirect "campania_email_configurar.asp?CampaniaID=" & CampaniaID & "&step=6"
    End If
%>
<h2>Paso 5: Frecuencia de Envio</h2>
<form method="post">
    <div class="form-group">
        <label>Cantidad por dia:</label>
        <input type="text" name="batchsize" required>
    </div>
	 
 
    <div class="form-actions" style="text-align:center; margin-top:25px;">
        <button type="submit" name="action" value="guardarPaso5" class="btn-volver">Confirmar Cambios</button>
		<button type="button" onclick="window.location.href='campania_email_configurar.asp?CampaniaID=<%=CampaniaID%>&step=4'" class="btn-cancelar">
			Volver
		</button>		
    </div>
</form>
<%
End If
%>



<%
' PASO 6: Fin
If step="6" Then
    
%>
<h2>Paso 6: Campaña guardada correctamente</h2>
 
<%
End If
%>


</div>
</body>
</html>
