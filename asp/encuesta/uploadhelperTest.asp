<!--#INCLUDE FILE="_upload.asp"-->
<!DOCTYPE html>

<%
Function do_Upload(rel_Folder)
  Server.ScriptTimeout = 3600
'Create upload form
'Using Huge-ASP file upload
'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm 
  '{b}Set the upload ID for this form.
  'Progress bar window will receive the same ID.
  Form.UploadID = Request.QueryString("UploadID")'{/b}

  Form.SizeLimit = 10*1024*1024 '10MB

  Dim HTML, hResult
  Const fsCompletted  = 0
  Const fsSizeLimit   = &HD
  Const fsTimeOut     = &HE
  Const fsError       = &HA
  
  If Form.State > fsError Then 'Some error state. 
    If Form.State = fsSizeLimit Then 'Data size exceeds limit. 
      hResult = "Upload size (" & Request.TotalBytes/1024 & "B) exceeds limit (" & Form.SizeLimit/1024 & "kB)."
    ElseIf Form.State = fsTimeOut Then 'Request timeout 
      hResult = "Upload time exceeds limit (" & Form.ReadTimeout & "s)."
    Else
      hResult = "Another upload problem (code " & Form.State & ")"
    End If
    hResult = "<Font Color=Red>" & hResult & "</Font><br>"
    Response.Status = "400 Bad request"
  ElseIf Form.State = fsCompletted Then 'Completted
    Form.Files.Save MapFolderToDisk(rel_Folder)
    hResult = "Files was saved to " & rel_Folder & " folder."
    hResult = "<Font Color=Green>" & hResult & "</Font><br>"
  ElseIf Request.QueryString("Action") = "Cancel" Then   
    hResult = "Upload was cancelled."  
  End If 

  '{b}get an unique upload ID for this upload script and progress bar.
  Dim UploadID, PostURL
  UploadID = Form.NewUploadID

  HTML = HTML & hResult
  HTML = HTML & "Upload files to '" & rel_Folder & "' folder. Limit:" & Form.SizeLimit/1024 & "kB.<br>"
  HTML = HTML & "<form method=post ENCTYPE=multipart/form-data Action=" & ref("UploadID=" & UploadID & "&Action=UPLOAD") & " OnSubmit=""return ProgressBar();"">"
  HTML = HTML & "<Input Type=Button Value=""+Add a file"" OnClick=Expand()>  <input type=submit value=""Upload files &gt;&gt;""><br>"
  HTML = HTML & "<Div ID=files> File 1 : <input type=file name=File1></Div>"
  HTML = HTML & "</Form>"
  HTML = HTML & "<"+"Script>var nfiles = 1;"
  HTML = HTML & "function Expand(){"
  HTML = HTML & "nfiles++;"
  HTML = HTML & "files.insertAdjacentHTML('BeforeEnd','<BR>File '+nfiles+' : <input type=file name=File'+nfiles+'>');"
  HTML = HTML & "};"
  HTML = HTML & "function ProgressBar(){" & vbCrLf
  HTML = HTML & "  var ProgressURL;" & vbCrLf
  HTML = HTML & "  ProgressURL = 'progress.asp?UploadID=" & UploadID & "'" & vbCrLf
  HTML = HTML & "  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200')" & vbCrLf
  HTML = HTML & "  return true;" & vbCrLf
  HTML = HTML & "};" & vbCrLf
  HTML = HTML & "</"+"Script>"
  HTML = HTML & ""
  do_Upload = HTML
End Function
%>

<html lang="es">
  <head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title></title>
	<script type="text/javascript"language="j78avascript"></script>	
  </head>
  <body>
<%=do_Upload("")%>

  
<!--#INCLUDE FILE="_common.asp"--> 
 

  	
  </body>
</html>