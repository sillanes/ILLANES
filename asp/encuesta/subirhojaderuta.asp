<%@ Language=VBScript %> 

<!--#INCLUDE FILE="_upload.asp"-->
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 
 

Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor","exec","execute",_
				  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
'Note: We can include following keyword to make a stronger scan but it will also 
'protect users to input these words even those are valid input
'  "!", "char", "alter", "begin", "cast", "create",  
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
 

For Each s in Request.Form 
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		' Redirect to an error page
		Response.Redirect(ErrorPage)
	End If
Next
%><!--#include virtual="./includes/sql-check.asp"--><%

 
If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If

doWhat = objRequest("doWhat")  
fileupload = objRequest("cmd")
subirarchivo = objRequest("subirarchivo")

if doWhat="" or doWhat<"3" Then
	Session("FileUploaded") = ""
End If

'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If
 
If doWhat = "-1"  Then 
	response.redirect "../menu.asp"
End If
  
If doWhat = "2"  Then 
	response.redirect "../menu.asp"
End If
  
  
 
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
      ' hResult = "Another upload problem (code " & Form.State & ")"
    End If
    hResult = "<Font Color=Red>" & hResult & "</Font><br>"
    Response.Status = "400 Bad request"
  ElseIf Form.State = fsCompletted Then 'Completted
    Form.Files.Save MapFolderToDisk(rel_Folder)
    hResult = "<div><Font Color=Green>" & hResult & "</Font><br></div>"
  ElseIf Request.QueryString("Action") = "Cancel" Then   
    hResult = "Upload was cancelled."  
  End If 

  '{b}get an unique upload ID for this upload script and progress bar.
  Dim UploadID, PostURL, Comment
  UploadID = Form.NewUploadID


  HTML = HTML & hResult
  
  HTML = HTML & "<form name=""FFUP"" method=post ENCTYPE=multipart/form-data Action=" & ref("UploadID=" & UploadID & "&Action=UPLOAD&doWhat=3") & " OnSubmit=""return ProgressBar();"">"
  HTML = HTML & "</br>"
  HTML = HTML & "<Div class=""pagetitle"" ID=files>Subir nuevo archivo: <input type=file name=File1></Div>"
  HTML = HTML & "<input type=submit value=""Subir Archivo""><br>"
  
  HTML = HTML & "</Form>"
  HTML = HTML & "<"+"Script>var nfiles = 1;"
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
  
Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

Set objRequest = Nothing
%>
<HTML>
<HEAD>
<TITLE>ILLANES HNOS SRL</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />
<script type="text/javascript" src="../includes/calendar_cool.js"></script>
<script type="text/javascript" src="../includes/copy.js"></script>
    <style>
        li {
            cursor: pointer;
			margin: 15px 0;
        }
    </style>
 
  
</HEAD>
<body>


	<form name="FF" method="post" action="subirhojaderuta.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	
 	
	</Form>
	
 
	
    <div class="wrap" style="width: 70%" align="center">
	<h1>HOJAS DE RUTA</h1>
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
		<!--#INCLUDE FILE="_commonHojaDeRuta.asp"--> 
		</div>
	<%
	End if
	
	
	If doWhat="3" Then

	
		if Session("FileUploaded") <> "" then
			
	%>
		<!--#include file="./includes/HojaDeRuta_Del.asp" -->
		
		<!--#include file="./includes/hojaderutaxls.asp" -->
			
		<!--#include file="./includes/HojaDeRuta_Ins.asp" -->
		
		<span align="center" class="pagetitle"> 
			<div align="center">
				<span class="textSuccess">Archivo subido correctamente: <%=Session("FileUploaded")%> </span>
			</div> 
	
		</span>		
	<%
		Else
	%>

	<%
		End If
	End If
	%>
 
  
	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(FF)" style="width:80px;">
	</div>
	
<script language="javascript">
function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
}


function chkForm(Fm,prm){
	var msg  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;
  
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.submit(); 
	}
}
  

 
</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


