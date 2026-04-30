<!-- #include file="../security/keycloak_func.asp" -->
<%
Function CheckUserAccess(byVal fUserID,byVal fGroupID,byVal fApplicationName,byVal fKeyName)

	'CPI-532 - Verify permissions in array
	If IsToggleOn() = True And IsToggleAuthOn() = True Then
		Dim PermArr
		PermArr = Session("PermArr")
		Dim tmp
		tmp = 0
		For i = 0 to UBound(PermArr)
			If fApplicationName = PermArr(i) Then
				tmp = 1
				Exit For
			End If
		Next
		CheckUserAccess = tmp

	Else

		rem ******************************************************************************
		rem This part is the login without AurisWebManager.UserLogin component! if needed!
		Dim dbCon77
		Set dbCon77 = Server.CreateObject("ADODB.Connection")
		dbCon77.Open "DRIVER=SQL Server;SERVER=SQLREPORTING01;APP=Web;DATABASE=Web", "webuser", "wwwpass"
		Dim cM77
		Set cM77 = Server.CreateObject("ADODB.Command")
		cM77.ActiveConnection = dbCon77
		cM77.CommandText = "wsp_preUserPermissions_check"
		cM77.CommandType = 4'adCmdStoredProcedure
		cM77.Parameters.Append cM77.CreateParameter("PermissionID",3,3,,0)
		cM77.Parameters.Append cM77.CreateParameter("ApplicationName",200,1,40,fApplicationName)
		cM77.Parameters.Append cM77.CreateParameter("KeyName",200,1,40,fKeyName)
		cM77.Parameters.Append cM77.CreateParameter("UserID",3,1,,fUserID)
		cM77.Parameters.Append cM77.CreateParameter("GroupID",3,1,,fGroupID)
		cM77.Execute
		CheckUserAccess = cM77("PermissionID")
		Set cM77 = Nothing
		dbCon77.Close
		Set dbCon77 = Nothing
		rem ******************************************************************************
	End If
End Function

Function CheckSession(fApplicationName,fKeyName,fReferrer)

	If IsToggleOn() = True Then
		If IsSessionActive() = False Then
			Response.Redirect (GetKeycloakConfig("LogoutUrl"))
		End If
	Else
		If Session("OldPassword")="**" Then Response.Redirect "../pwcheck_forced.asp"
		If Session("pwcheck")="*" Then Response.Redirect "../pwcheck.asp"
		If CLng("0"&Session("WebcsOnly"))=1 Then Response.Redirect "../dummy.asp?code=4"
	End If

	Dim fUserID
	fUserID = Session("pre_userid")&""
	If fUserID = "" Then
		Response.Redirect "../dummy.asp?code=3"
	ElseIf fApplicationName<>"" Or fKeyName<>"" Then
		If CheckUserAccess(fUserID,0,fApplicationName,fKeyName) = 0 Then Response.Redirect "../dummy.asp?code=4"
	End If
End Function

Function ParseValueFromString (byVal fString,byVal fIndex,byVal fSeperator,byRef fValue)
	Dim a0, a1, a2, a3
	a0 = 1
	a1 = 0
	a2 = 0
	a3 = ""
	a1 = InStr(a0, fString, fSeperator)
	Do While a1 <> 0 and a2 < fIndex
		a2 = a2 + 1
		If a2 = fIndex Then
			a3 = Mid(fString,a0,a1-a0)
		Else
			a3 = ""
		End If
		a0 = a1 + 1
		a1 = InStr(a0, fString, fSeperator)
	Loop
	If a2 + 1 = fIndex and a1 = 0 Then
		a3 = Mid(fString,a0,len(fString) - a0 + 1)
	End If
	fValue = a3
End Function

Sub setCommandParameters(spSTR, cM)
	Dim i, pos1, pos2, pos3, spLen, token, ch, theRequestObj
	Dim paramName, paramType, paramDirection, paramSize, paramVal, paramSelecteds

	If Request.Form.Count > 0 Then
		Set theRequestObj = Request.Form
	Else
		Set theRequestObj = Request.QueryString
	End If
	paramSelecteds = theRequestObj("paramSelecteds")&""

	pos2  = 0
	spLen = Len(spSTR)
	While pos2 < spLen
		'Search for parameter name
		pos1 = pos2 + 1
		pos2 = InStr(pos1,spSTR,"@")
		pos1 = pos2 + 1
		pos2 = InStr(pos1,spSTR," ")
		pos3 = InStr(pos1,spSTR,chr(9))
		If pos2>pos3 and pos3<>0 Then pos2 = pos3
		paramName = Mid(spSTR,pos1,pos2-pos1)
		pos1 = pos2 + 1

		'Search for parameter type
		pos2 = InStr(pos1,spSTR,",")
		If pos2 = 0 Then pos2 = spLen
		token = ""
		For i = pos1 to pos2
			ch = LCase(Mid(spSTR,i,1))
			If asc(ch)>=97 And asc(ch)<=122 Then
				token = token & ch
			ElseIf token<>"" Then
				Exit For
			End If
		Next
		pos1=i

		'Set parameter direction
		paramDirection = 1
		pos2 = InStr(pos1,spSTR,"@")-1
		If pos2 < 0  Then pos2 = spLen
		pos3 = InStr(pos1,spSTR,"out",1)
		If pos3<>0 And pos3<pos2 Then paramDirection = 3

		'Set parameter value
		paramVal = Trim(theRequestObj(paramName)&"")

		'Set parameter type
		If token = "varchar" Then
			paramType = 200
			paramSize = Abs(CInt(Mid(spSTR,pos1,InStr(pos1,spSTR,")")-pos1+1)))
		ElseIf token = "int" or token = "integer" or token = "tinyint" or token = "smallint" Then
			paramType = 3
			If isNumeric(paramVal) Then paramVal = CLng(paramVal)
		ElseIf token = "text" Then
			paramType = 200
			paramSize = 1000
		ElseIf token = "bit" Then
			paramType = 11
			If isNumeric(paramVal) Then paramVal = CBool(paramVal)
		ElseIf token = "numeric" or token = "float" Then
			paramType = 5
			If isNumeric(paramVal) Then paramVal = CDbl(paramVal)
		ElseIf token = "smalldatetime" or token = "datetime" Then
			paramType = 135
			paramVal  = Replace(paramVal,".","/")
			If isDate(paramVal) Then paramVal = CDate(paramVal)
		End If

		If paramSelecteds<>"" Then
			If InStr(paramSelecteds,paramName)=0 Then paramVal = Null
		End If
		If paramType=200 Then
			If paramSendNull_Chr And paramVal = "" Then paramVal = Null
			cM.Parameters.Append cM.CreateParameter(paramName,paramType,paramDirection,paramSize,paramVal)
		Else
			If paramVal = "" Then paramVal = Null
			cM.Parameters.Append cM.CreateParameter(paramName,paramType,paramDirection,,paramVal)
		End If
  
		'response.write paramName&"-"&paramType&"-"&paramDirection&"-"&paramSize&"-"&paramVal&"<br>"
		'if err.number<>0	then
		'	response.write err.description&"<br>"
		'	response.end
		'end if
	Wend
End Sub

Function GetCmParams(cM, idx)
	Dim i, returnStr, objPrm

	returnStr = ""
	For i = idx to cM.Parameters.Count-1
		Set objPrm = cM.Parameters.Item(i)
		If isNull(objPrm.Value) Then
			returnStr = returnStr & ", @" & objPrm.Name & "=Null"
		ElseIf objPrm.Type = 200 or objPrm.Type = 135 Then
			returnStr = returnStr & ", @" & objPrm.Name & "='" & objPrm.Value & "'"
		Else
			returnStr = returnStr & ", @" & objPrm.Name & "=" & objPrm.Value
		End If
	Next

	returnStr = Mid(returnStr,3)
	If idx = 0 Then
		i = InStr(cM.CommandText," call ")
		GetCmParams = "EXEC " & Mid(cM.CommandText,i+6,InStr(i+6,cM.CommandText,"(?")-i-6) & " " & returnStr
	Else
		GetCmParams = returnStr
	End If
End Function

Sub LogUserAction(dbCon, CompanyID, ProductID, ActionSection, ActionText)
	Dim cM2
	Set cM2 = Server.CreateObject("ADODB.Command")
	cM2.ActiveConnection = dbCon
	cM2.CommandText = "wsp_preUserActionHistory_ins"
	cM2.CommandType = 4'adCmdStoredProcedure
	cM2.Parameters.Append cM2.CreateParameter("CompanyID",3,1,,CLng("0" & CompanyID))
	cM2.Parameters.Append cM2.CreateParameter("ProductID",3,1,,CLng("0" & ProductID))
	cM2.Parameters.Append cM2.CreateParameter("UserID",3,1,,CLng("0"&Session("pre_userid")))
	cM2.Parameters.Append cM2.CreateParameter("RemoteIP",200,1,20,Left(Request.ServerVariables("REMOTE_ADDR"),20))
	cM2.Parameters.Append cM2.CreateParameter("Section",200,1,20,ActionSection)
	cM2.Parameters.Append cM2.CreateParameter("ActionText",200,1,2000,Left(ActionText,2000))
	cM2.Execute
	Set cM2 = Nothing
End Sub


Sub LogEndUserAction(dbCon, CompanyID, ProductID, CardID, AccountID, PIN, ANI, ActionSection, ActionText, ActionDescription, Username)
	Dim cM2
	Set cM2 = Server.CreateObject("ADODB.Command")
	cM2.ActiveConnection = dbCon
	cM2.CommandText = "wsp_EndUserActionHistory_ins"
	cM2.CommandType = 4'adCmdStoredProcedure
	cM2.Parameters.Append cM2.CreateParameter("UserID",3,1,,CLng("0"&Session("pre_userid")))
	cM2.Parameters.Append cM2.CreateParameter("CompanyID",3,1,,CLng("0" & CompanyID))
	cM2.Parameters.Append cM2.CreateParameter("ProductID",3,1,,CLng("0" & ProductID))
	cM2.Parameters.Append cM2.CreateParameter("CardID",3,1,,CLng("0" & CardID))
	cM2.Parameters.Append cM2.CreateParameter("AccountID",20,1,,CLng("0" & AccountID))
	cM2.Parameters.Append cM2.CreateParameter("PIN",200,1,10,PIN)
	cM2.Parameters.Append cM2.CreateParameter("ANI",200,1,28,ANI)
	cM2.Parameters.Append cM2.CreateParameter("Section",200,1,20,ActionSection)
	cM2.Parameters.Append cM2.CreateParameter("ActionText",200,1,2000,Left(ActionText,2000))
	cM2.Parameters.Append cM2.CreateParameter("ActionDescription",200,1,2000,Left(ActionDescription,2000))
	cM2.Parameters.Append cM2.CreateParameter("RemoteIP",200,1,20,Left(Request.ServerVariables("REMOTE_ADDR"),20))
	cM2.Parameters.Append cM2.CreateParameter("Username",200,1,20,Username)
	cM2.Execute
	Set cM2 = Nothing
End Sub


Function FormatNumeric(numType, numVal, afterDecimal)
	If numVal<>"" Then
		If numType="C" Then
			If CTAG<>"" Then
				FormatNumeric = Replace(FormatCurrency(CDbl(numVal)*CDbl(CCON),afterDecimal),"$",Left(CTAG,1))
			Else
				FormatNumeric = FormatCurrency(numVal,afterDecimal)
			End If
			If InStr(FormatNumeric,"(") > 0 Then FormatNumeric = "<font color='#ff0000'>" & "-" & FormatNumeric & "</font>"
		ElseIf numType="N" Then
			FormatNumeric = FormatNumber(numVal,afterDecimal)
			If InStr(FormatNumeric,"-") > 0 Then FormatNumeric = "<font color='#ff0000'>" & FormatNumeric & "</font>"
		End If
	Else
		FormatNumeric = ""
	End If
End Function

Function FormatNumericZ(numType, numVal, afterDecimal)
	If numVal="" Then numVal="0"

	If numVal<>"" Then
		If numType="C" Then
			If CTAG<>"" Then
				FormatNumericZ = Replace(FormatCurrency(CDbl(numVal)*CDbl(CCON),afterDecimal),"$",Left(CTAG,1))
			Else
				FormatNumericZ = FormatCurrency(numVal,afterDecimal)
			End If
			If InStr(FormatNumericZ,"(") > 0 Then FormatNumericZ = "<font color='#ff0000'>" & "-" & FormatNumericZ & "</font>"
		ElseIf numType="N" Then
			FormatNumericZ = FormatNumber(numVal,afterDecimal)
			If InStr(FormatNumericZ,"-") > 0 Then FormatNumericZ = "<font color='#ff0000'>" & FormatNumericZ & "</font>"
		End If
	End If
End Function

Function FormatIndex(numVal)
	Randomize
	FormatIndex = Left(Replace((Rnd()*123456)&"",".",""),8) & numVal & Left(Replace((Rnd()*123456)&"",".",""),8)
End Function

'' Added by Carlos Torres on 10/15/2018 - Basic function to send text email
Sub SendMessage(FromAddress, ToAddress, Subject, TextMessage)

	Dim objEmail


	Set objEmail = Server.CreateObject("CDO.Message")

	' Set Email Headers
	objEmail.From = FromAddress
	objEmail.To = ToAddress
	objEmail.Subject = Subject
	' Construct Email Body
	objEmail.textbody = TextMessage

	' Configure email server
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mx2.anewbroadband.com"
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
	objEmail.Configuration.Fields.Update

	' send email
	objEmail.Send

End Sub

'' Added by Carlos Torres on 4/12/2022 - Basic function to send html body email
Sub SendHTMLMessage(FromAddress, ToAddress, Subject, TextMessage)

	Dim objEmail


	Set objEmail = Server.CreateObject("CDO.Message")

	' Set Email Headers
	objEmail.From = FromAddress
	objEmail.To = ToAddress
	objEmail.Subject = Subject
	' Construct Email Body
	objEmail.htmlbody = "<html><body>" & TextMessage & "</body></html>"

	' Configure email server
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mx2.anewbroadband.com"
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
	objEmail.Configuration.Fields.Update

	' send email
	objEmail.Send

End Sub

''Added by Reynaldo Lazaro on 2/18/2022  -  Encode and Decode general functios
Function StrToHex(value)
	Dim result, i, c
	
	result = ""
	For i = 1 To Len(value)
		c = CStr(Hex(Asc(Mid(value, i, 1))))
		result = result & c
	Next
	StrToHex = result	
End Function

Function HexToStr(value)
	Dim result, i, c
	
	result = ""
	For i = 1 To Len(value) Step 2
		c = CLng("&H" & Mid(value, i, 2))
		result = result & Chr(c)
	Next
	HexToStr = result
End Function

Function Base64Encode(sText)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue =Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

Function IsValidANI(ani)
  Dim checkType
  Set checkType = New RegExp
  checkType.Pattern = "[^0-9]"
      
  if Len(ani) < 8 or checkType.Test(ani) = True then 
    IsValidANI = False 
  elseIf Left(ani,1) = "0" and Left(ani,3) <> "011" then
    IsValidANI = False
  else 
    IsValidANI = True
  end if
End Function

%><!--[if IE 8]><!DOCTYPE html><![endif]-->