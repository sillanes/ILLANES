<%
CookieExceptionList = Array("""","(",")")

Function IsExceptionList(str,varType)
    If(varType="cookie") then
        For Each item in CookieExceptionList
            If(item=str) then
                IsExceptionList=True
                Exit Function
            End If
        Next
    End If
    IsExceptionList=False
End Function


 
Function CheckStringForSQL(str,varType) 
  On Error Resume Next 
  Dim lstr 
  ' If the string is empty, return false that means pass
  If ( IsEmpty(str) ) Then
    CheckStringForSQL = false
    Exit Function
  ElseIf ( StrComp(str, "") = 0 ) Then
    CheckStringForSQL = false
    Exit Function
  End If
  
  lstr = LCase(str)
  ' Check if the string contains any patterns in our black list
  For Each s in BlackList
    If(IsExceptionList(s,varType)=False) then
        If ( InStr (lstr, s) <> 0 ) Then
          CheckStringForSQL = true
          Exit Function
        End If
    End If
  Next
  CheckStringForSQL = false
End Function 

  
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Check Upload forms data
'  Description: This function will validate ASP Upload Data
'  Note:        Because of ASPUpload's limitation this function 
'               need to be called after its save function from 
'               the relevant ASP page
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function IsValidUploadFormData(dataCollection,redirect)
    for each item in dataCollection
        If ( CheckStringForSQL(item) ) Then
            PrepareReport("Upload Form")
            'Redirect to an error page
            if(redirect) then Response.Redirect(ErrorPage)
            IsValidUploadFormData = false
            Exit Function
         End If
    next
    IsValidUploadFormData = true
end function

Function PrepareReport(injectionType)
    'Build the messege
    Dim MessageBody
    MessageBody="<HTML><HEAD></HEAD><BODY><h1>One Sql Injection Attempt Was Blocked! </h1><br/>"
    MessageBody=MessageBody & "Attack Time: " & FormatDateTime(Now,3) & "<br/>"
    MessageBody=MessageBody & "Attaker IP Address: " &  Request.ServerVariables("REMOTE_HOST") & "<br/>"
    MessageBody=MessageBody & "Injection Type: " & injectionType & "<hr size='1'/><br/>"
    MessageBody=MessageBody & "More Details Information: <br/>"    
    MessageBody=MessageBody&"<table width='100%'>"
    MessageBody=MessageBody& "<tr><td colspan='2'><h2>Form Variables</h2></td></tr>"
    'List Form Data
    For Each s in Request.Form
        MessageBody=MessageBody&"<tr>"
        MessageBody=MessageBody&"   <td>" & s & "</td>"
        MessageBody=MessageBody&"   <td>" & Request.Form(s) & "</td>"
        MessageBody=MessageBody&"<tr>"
    Next
    MessageBody=MessageBody&   "<tr><td colspan='2'><h2>QueryString Variables</h2></td></tr>"
    For Each s in Request.QueryString
        MessageBody=MessageBody&"<tr>"
        MessageBody=MessageBody&"   <td>" & s & "</td>"
        MessageBody=MessageBody&"   <td>" & Request.QueryString(s) & "</td>"
        MessageBody=MessageBody&"<tr>"
    Next
 
    MessageBody=MessageBody&    "<tr><td colspan='2'><h2>Cookie Variables</h2></td></tr>"
    For Each s in Request.Cookies
        MessageBody=MessageBody&"<tr>"
        MessageBody=MessageBody&"   <td>" & s & "</td>"
        MessageBody=MessageBody&"   <td>" & Request.Cookies(s) & "</td>"
        MessageBody=MessageBody&"<tr>"
    Next
    
    MessageBody=MessageBody&"</table><br/>"
    MessageBody=MessageBody & "Script Page: " & GetCurrentPageUrl() & "<br/>"
    MessageBody=MessageBody & "Referer Page: " & GetRefererPageUrl() &  "<br/><br/>"
    MessageBody=MessageBody & "</BODY></HTML>"
    Result= SendEmail("Sql Injection Attempt Was Detected by " & injectionType & "!",MessageBody)
End Function
Function GetCurrentPageUrl()
    domainname = GetCurrentServerName() 
    filename = Request.ServerVariables("SCRIPT_NAME") 
    querystring = Request.ServerVariables("QUERY_STRING") 
    GetCurrentPageUrl= domainname & filename & "?" & querystring 
End Function
 
Function GetRefererPageUrl()
    GetRefererPageUrl= Request.ServerVariables("HTTP_REFERER") 
End Function
 
Function GetCurrentServerName()
    prot = "http" 
    https = lcase(request.ServerVariables("HTTPS")) 
    if https <> "off" then prot = "https" 
    domainname = Request.ServerVariables("SERVER_NAME") 
    GetCurrentServerName=prot & "://" & domainname 
End Function
 
Function GetPageNameFromPath(strPath)
    strPos= len(strPath)-InStrRev(strPath,"/")
    pageName=right(strPath,strPos)
    GetPageNameFromPath=pageName
End Function
 
Function GetCurrentPageName()
    scriptPath = Request.ServerVariables("SCRIPT_NAME") 
    pageName=GetPageNameFromPath(scriptPath)
    GetCurrentPageName=pageName
End Function

Function SendEmail(subject,html)
	SendEmail = "ERROR"
	Set emailObj      = CreateObject("CDO.Message")

	emailObj.From     = "reclamosweb@illanes.com.ar"
	emailObj.To       = "sillanes@illanes.com.ar"

	emailObj.Subject  = subject
	emailObj.HTMLBody = html

	Set emailConfig = emailObj.Configuration

	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.illanes.com.ar"
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2  
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
	'emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")      = true 
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")    = "reclamosweb@illanes.com.ar"
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")    = "Illanes%Reclamo$2023+"

	emailConfig.Fields.Update
	If err.number = 0 Then Response.write  err
	emailObj.Send

	If err.number = 0 then SendEmail =  "OK"
End Function
%>