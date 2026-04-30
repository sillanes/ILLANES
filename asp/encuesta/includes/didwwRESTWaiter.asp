<%
Function ddwwCancelDID(didww_did_number)
	Dim ddwwURL, ddwwResponse, xmlDoc, xmlParentNode
	Dim didww_errormessage, didww_result, didww_error
	
	On Error Resume Next
	
	didww_error = ""
	didww_result = ""
	didww_errormessage = ""
	
	'ddwwURL = "http://waiter.auris.com/DIDService.aspx?provider=DIDWW&provider_version=CIMA_Test&method=cancel&user=rey&ukey=ReyTest01&did_number=" & didww_did_number&""
	'ddwwURL = "http://192.168.166.31/DIDService.aspx?provider=RESTDIDWW&method=cancel&user=rey&ukey=ReyTest01&did_number=" & didww_did_number&""
	ddwwURL = "http://192.168.166.31/DIDService.aspx?provider=RESTDIDWW&method=cancel&user=rey&ukey=ReyTest01&did=" & didww_did_number&""
	'LogMessage("[ddwwCancelDID] Inputs: didww_did_number=" & didww_did_number)
	LogMessage_Test("[ddwwGroup] URL = " & ddwwURL)
	Set xmlDoc = Server.CreateObject("Msxml2.DOMDocument.6.0")
	xmlDoc.async = False
	
	ddwwResponse = ddwwSendRequest(ddwwURL)
	'LogMessage("[ddwwCancelDID] Response from DIDWW: " & ddwwResponse)
	LogMessage_Test("[ddwwCancelDID] Response from DIDWW: " & ddwwResponse)
	xmlDoc.LoadXml(ddwwResponse)
	If xmlDoc.hasChildNodes = True Then

		Set xmlParentNode = xmlDoc.selectSingleNode("/root/results")

		If IsObject(xmlParentNode) AND xmlParentNode.hasChildNodes = True Then

			For Each xmlChildNode In xmlParentNode.ChildNodes

				If xmlChildNode.nodeName = "errormessage" Then
					didww_errormessage = xmlChildNode.Text
				End If

				If xmlChildNode.nodeName = "result" Then
					didww_result= xmlChildNode.Text
				End If

				If xmlChildNode.nodeName = "error" Then
					didww_error = xmlChildNode.Text
				End If

			Next
		End If
	End If	
	
	If didww_result = "0" Then
		ddwwCancelDID = "OK"
	Else
		ddwwCancelDID = "ERROR:9302:Internal Error. Provider failed to cancel DID."
	End If
		
End Function

Function ddwwGroup(group_id)
	Dim ddwwURL, ddwwResponse, xmlDoc, cityNode, xmlChildNode, cityNodesList, errorMessage
	Dim strResponse, city_prefix, isAvailable, isLnrRequired
	
	strResponse = ""
	On Error Resume Next
	
	'ddwwURL = "http://waiter.auris.com/DIDService.aspx?provider=DIDWW&provider_version=CIMA_Test&method=order&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&""
	'ddwwURL = "http://192.168.166.31/test/DIDService.aspx?provider=RESTDIDWW&method=group&user=rey&ukey=ReyTest01&group_id="&group_id&""
	ddwwURL = "http://192.168.166.31/DIDService.aspx?provider=RESTDIDWW&method=group&user=rey&ukey=ReyTest01&group_id="&group_id&""
	
	LogMessage_Test("[ddwwGroup] URL = " & ddwwURL)
	LogMessage_Test("[ddwwGroup] Inputs: group_id="&group_id&"")
	ddwwResponse = ddwwSendRequest(ddwwURL)
	
	If err.Number <> 0 Then		
		ddwwResponse = "<root><results><errormessage>Error</errormessage><regions></regions></results></root>"
		LogMessage_Test("[ddwwCountryRegion] Error :: " & err.Number&"")
	End If
	
	LogMessage_Test("[ddwwGroup] Response from DIDWW: " & ddwwResponse)	
	
	ddwwGroup = ddwwResponse
	
End Function

Function ddwwCountryRegion(didww_country_iso, didww_city_prefix)
	Dim ddwwURL, ddwwResponse, xmlDoc, cityNode, xmlChildNode, cityNodesList, errorMessage
	Dim strResponse, city_prefix, isAvailable, isLnrRequired
	
	strResponse = ""
	On Error Resume Next
	
	'ddwwURL = "http://waiter.auris.com/DIDService.aspx?provider=DIDWW&provider_version=CIMA_Test&method=order&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&""
	'ddwwURL = "http://192.168.166.31/test/DIDService.aspx?provider=RESTDIDWW&method=regions&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&city_prefix="&didww_city_prefix&""
	ddwwURL = "http://192.168.166.31/DIDService.aspx?provider=RESTDIDWW&method=regions&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&city_prefix="&didww_city_prefix&""
	
	LogMessage_Test("[ddwwCountryRegion] URL = " & ddwwURL)
	LogMessage_Test("[ddwwCountryRegion] Inputs: didww_country_iso=" & didww_country_iso & "; didww_city_prefix=" & didww_city_prefix)
	ddwwResponse = ddwwSendRequest(ddwwURL)
		
	LogMessage_Test("[ddwwCountryRegion] Response from DIDWW: " & ddwwResponse)
	
	'Set xmlDoc = Server.CreateObject("Msxml2.DOMDocument.6.0")
	'xmlDoc.async = False
	'xmlDoc.LoadXml(ddwwResponse)
	
	If err.Number <> 0 Then		
		ddwwResponse = "<root><results><errormessage>Error</errormessage><regions></regions></results></root>"
		LogMessage_Test("[ddwwCountryRegion] Error :: " & err.Number&"")
	End If
	
	ddwwCountryRegion = ddwwResponse
	
End Function

Function ddwwCountryRegions(didww_country_iso)
	Dim ddwwURL, ddwwResponse, xmlDoc, cityNode, xmlChildNode, cityNodesList, errorMessage
	Dim strResponse, city_prefix, isAvailable, isLnrRequired
	
	strResponse = ""
	On Error Resume Next
	
	'ddwwURL = "http://waiter.auris.com/DIDService.aspx?provider=DIDWW&provider_version=CIMA_Test&method=order&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&""
	'ddwwURL = "http://192.168.166.31/test/DIDService.aspx?provider=RESTDIDWW&method=regions&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&restriction_allowed=1"&""
	ddwwURL = "http://192.168.166.31/DIDService.aspx?provider=RESTDIDWW&method=regions&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&restriction_allowed=1"&""
		
	LogMessage_Test("[ddwwCountryRegions] URL = " & ddwwURL)
	LogMessage_Test("[ddwwCountryRegions] Inputs: didww_country_iso=" & didww_country_iso)
	ddwwResponse = ddwwSendRequest(ddwwURL)
	
	If err.Number <> 0 Then		
		ddwwResponse = "<root><results><errormessage>Error</errormessage><regions></regions></results></root>"
		LogMessage_Test("[ddwwCountryRegion] Error :: " & err.Number&"")
	End If
		
	LogMessage_Test("[ddwwCountryRegions] Response from DIDWW: " & ddwwResponse)
	
	ddwwCountryRegions = ddwwResponse				
	
End Function

Function ddwwOrderDid(didww_country_iso, didww_city_prefix, sCountryID, order_id)
	Dim ddwwURL, ddwwResponse, sOrderReference, sRegionName, didww_order_id, xmlDoc, xmlParentNode, xmlChildNode
	Dim didww_country_name, didww_city_name, didww_did_number, didww_errormessage
		
	On Error Resume Next
	didww_country_name = ""
	sOrderReference = ""		
	didww_city_name = ""
	didww_did_number = ""
	sRegionName = ""
	didww_order_id = ""
	
	'ddwwURL = "http://waiter.auris.com/DIDService.aspx?provider=DIDWW&provider_version=CIMA_Test&method=order&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&city_prefix="&didww_city_prefix&""
	'ddwwURL = "http://192.168.166.31/test/DIDService.aspx?provider=RESTDIDWW&method=order&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&city_prefix="&didww_city_prefix&""
	ddwwURL = "http://192.168.166.31/DIDService.aspx?provider=RESTDIDWW&method=order&user=rey&ukey=ReyTest01&country_iso="&didww_country_iso&"&city_prefix="&didww_city_prefix&""
	If order_id <> "" Then
		ddwwURL = ddwwURL + "&sku_id="&order_id&""
	End If
		
	'LogMessage("[ddwwOrderDid] Inputs: didww_country_iso=" & didww_country_iso & "; didww_city_prefix=" & didww_city_prefix & "; CountryID=" & sCountryID)
	LogMessage_Test("[ddwwOrderDid] URL = " & ddwwURL)
	LogMessage_Test("[ddwwOrderDid] Inputs: didww_country_iso=" & didww_country_iso & "; didww_city_prefix=" & didww_city_prefix & "; CountryID=" & sCountryID)
	
	Set xmlDoc = Server.CreateObject("Msxml2.DOMDocument.6.0")
	xmlDoc.async = False
	
	ddwwResponse = ddwwSendRequest(ddwwURL)
	
	'LogMessage("[ddwwOrderDid] Response from DIDWW: " & ddwwResponse)
	LogMessage_Test("[ddwwOrderDid] Response from DIDWW: " & ddwwResponse)

	xmlDoc.LoadXml(ddwwResponse)

	If xmlDoc.hasChildNodes = True Then

		Set xmlParentNode = xmlDoc.selectSingleNode("/root/results")

		If IsObject(xmlParentNode) AND xmlParentNode.hasChildNodes = True Then

			For Each xmlChildNode In xmlParentNode.ChildNodes
			
				If xmlChildNode.nodeName = "errormessage" Then
					didww_errormessage = Trim(xmlChildNode.Text)
				End If			
				If xmlChildNode.nodeName = "country_name" Then
					didww_country_name = xmlChildNode.Text
				End If
				If xmlChildNode.nodeName = "area_name" Then
					didww_city_name = xmlChildNode.Text
				End If				
				If xmlChildNode.nodeName = "did_number" Then
					didww_did_number = xmlChildNode.Text
				End If				
				If xmlChildNode.nodeName = "order_id" Then
					didww_order_id = xmlChildNode.Text
				End If
			Next
		End If
	End If
	
	If didww_did_number <> "" Then
		If sCountryID = "1" Or sCountryID = "901" Then ' USA + Puerto Rico
			didww_did_number = Right(didww_did_number,10)
			sRegionName = didww_city_name
		ElseIf sCountryID = "39" Or sCountryID = "61" Then ' CANADA + DOMINICAN REP
			didww_did_number = Right(didww_did_number,10)
			sRegionName = didww_country_name & "-" & didww_city_name
		Else
			didww_did_number = "011" & didww_did_number
			sRegionName = didww_country_name & "-" & didww_city_name
		End If
		LogMessage_Test("[ddwwOrderDid] Response : " & didww_did_number & ":" & didww_order_id & ":" & sRegionName)
		ddwwOrderDid = didww_did_number & ":" & didww_order_id & ":" & sRegionName
		
	ElseIf didww_errormessage <> "" Then
		ddwwOrderDid = "ERROR:9305:Provider Message. " & didww_errormessage
	Else
		ddwwOrderDid = "ERROR:9305:Internal Error. Failed to provision an order with the provider."
	End If
	
End Function

Function ddwwSendRequest(strURL)
	Dim dwwHTTP, dwwResponse
	dwwResponse = ""
	On Error Resume Next	
	Set dwwHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	dwwHTTP.setTimeouts (50 * 1000), (50 * 1000), (150 * 1000), (150 * 1000)
	
	If (err.Number = 0) AND (dwwHTTP.readyState = 0) Then
		dwwHTTP.Open "GET", strURL, false
	End If
	If (err.Number = 0) AND (dwwHTTP.readyState = 1) Then
		dwwHTTP.Send
	End If
	If (err.Number = 0) AND (dwwHTTP.readyState = 4) Then
		strResponse = dwwHTTP.responseText&"" 'Mid(Trim(dwwHTTP.responseText&""),1,12000)
	End If
	ddwwSendRequest = strResponse
End Function

Private Sub LogMessage_Test(sMessage) 
	
	Dim fso, file, today
	
	''Response.Write("LOG: " & sMessage & "<br><br>" & vbCrLf & vbCrLf)
	On Error Resume Next
	today = CStr(Year(Now)) & Right("0"&CStr(Month(Now)),2) & Right("0"&CStr(Day(Now)),2)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile("C:\Temp\logs\TestDIDWW_" & today & ".log", 8, true)
	file.Write "[" & CStr(Session.SessionID) & "] "
	file.Write "[" & FormatDateTime(Now,0) & "] "
	file.Write "[didwwRESTWaiter]"
	file.WriteLine sMessage
	file.Close
	
	Set file = Nothing
	Set fso = Nothing
	
End Sub
%>