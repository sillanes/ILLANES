<%
' Implemented by Carlos Torres
' Date 8/11/2021
' Description: Functions to connect to all the Auris portals, to validate
' if given usernames are already on use accross different portals.
Function PreValidateUsers( aUsers )

	Dim portals, portal
	Dim ConnList
	Dim oConn, oCmd, iRows
	Dim idx, sUser, sPortal, aUser, dUsers, Counter

	Set PreValidateUsers = Nothing

	If IsArray(aUsers) Then ' must come already as a array of users from the Form (4 columns on each user)
		If UBound(aUsers) >= 0 Then ' One element array, ubound will return 0

			' Preload users that meet requirements: 4 columns and username no longer than 10 characters
			Counter = 0
			Set dUsers = Server.CreateObject("Scripting.Dictionary")
			For Each sUser In aUsers
				sUser = Trim(sUser&"")
				aUser = Split(sUser, ",")		
				If UBound(aUser) = 3 Then
					If Len(Trim(aUser(1)&"")) <= 10 Then
						dUsers.Add Counter, UCase(Trim(aUser(1)&""))
						Counter = Counter + 1
					End If
				End If
			Next

			If dUsers.Count > 0 Then
				portals = Split("wholesale,retail,tracfone,rtfd,topup",",")
				Set UserList = Server.CreateObject("Scripting.Dictionary")
				Set ConnList = Server.CreateObject("Scripting.Dictionary")
				LoadConnInfo( ConnList )

				If ConnList.Count > 0 Then  
					Set oConn = Server.CreateObject("ADODB.Connection")
					For Each portal In portals
						''Response.Write("<li>Checking portal..." & portal & "</li>")
						iRows = 0
						oConn.Open  ConnList.Item(portal).ConnStr, ConnList.Item(portal).ConnUsr, ConnList.Item(portal).ConnPwd
						Set oCmd = Server.CreateObject("ADODB.Command")
						oCmd.ActiveConnection = oConn
						oCmd.CommandText = ConnList.Item(portal).StorProc
						oCmd.CommandType = adCmdStoredProcedure
						oCmd.Parameters.Append oCmd.CreateParameter("UserName",adParamVarchar,1,50,"")
						oCmd.Parameters.Append oCmd.CreateParameter("RowsAffected",adParamInt,3,,0)
						
						For Each sUser In dUsers.Items
							oCmd.Parameters.Item("UserName").Value = sUser
							oCmd.Execute
							iRows = CLng(oCmd("RowsAffected"))
							If iRows > 0 Then
								If UserList.Exists(sUser) Then
									UserList.Item(sUser) = UserList.Item(sUser) & "," & portal
								Else
									UserList.Add sUser, portal
								End If
							End If
							''Response.Write("<li>" & sUser & " - " & UserList.Item(sUser) & "</li>")
							
						Next ' user loop
						Set oCmd = Nothing
						oConn.Close						
					Next  ' portal loop
					Set oConn = Nothing
				End If

				If UserList.Count > 0 Then
					Set PreValidateUsers = UserList
				End If
			End If
			
			Set dUsers = Nothing
			Set ConnList = Nothing	
			Set UserList = Nothing
		
		End If	 ' aUers must have at least 1 member
	End If       ' aUsers must be an array
	
End Function

Sub LoadConnInfo(ByRef dList )

	'' Loading Wholesale
	Set oConnInfo = New ConnInfo
	oConnInfo.ConnStr = "DRIVER=SQL Server;SERVER=awsqlswi01;APP=Web;DATABASE=Wholesale"
	oConnInfo.ConnUsr = "webuser"
	oConnInfo.ConnPwd = "wwwpass"
	oConnInfo.StorProc = "users.wsp_user_exists"
	dList.Add "wholesale", oConnInfo
	Set oConnInfo = Nothing
	'' Loading Retail
	Set oConnInfo = New ConnInfo
	oConnInfo.ConnStr = "DRIVER=SQL Server;SERVER=SQLREPORTING01;APP=Web;DATABASE=Web"
	oConnInfo.ConnUsr = "webuser"
	oConnInfo.ConnPwd = "wwwpass"
	oConnInfo.StorProc = "wsp_preUsers_exists"
	dList.Add "retail", oConnInfo
	Set oConnInfo = Nothing
	'' Loading Topup
	Set oConnInfo = New ConnInfo
	oConnInfo.ConnStr = "DRIVER=SQL Server;SERVER=SQLSWITCH;APP=Web;DATABASE=txnbilling"
	oConnInfo.ConnUsr = "webuser"
	oConnInfo.ConnPwd = "wwwpass"
	oConnInfo.StorProc = "users.wsp_User_exists"
	dList.Add "topup", oConnInfo
	Set oConnInfo = Nothing	
	'' Loading RTFD
	Set oConnInfo = New ConnInfo
	oConnInfo.ConnStr = "DRIVER=SQL Server;SERVER=SQLREPORTING01;APP=Web;DATABASE=fraud"
	oConnInfo.ConnUsr = "webuser"
	oConnInfo.ConnPwd = "wwwpass"
	oConnInfo.StorProc = "dbo.wsp_Users_exists"
	dList.Add "rtfd", oConnInfo
	Set oConnInfo = Nothing
	'' Loading Tracfone
	Set oConnInfo = New ConnInfo
	oConnInfo.ConnStr = "DRIVER=SQL Server;SERVER=SQLREPORTING01;APP=Web;DATABASE=WebTracfone"
	oConnInfo.ConnUsr = "webuser"
	oConnInfo.ConnPwd = "wwwpass"
	oConnInfo.StorProc = "wsp_preUsers_exists"
	dList.Add "tracfone", oConnInfo
	Set oConnInfo = Nothing

End Sub

Class ConnInfo

  Public ConnStr
  Public ConnUsr
  Public ConnPwd
  Public StorProc

End Class

%>