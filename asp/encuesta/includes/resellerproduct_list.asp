<%
sSQL = "EXEC wsp_CompanyProducts_Sel " & CLng("0"&Reseller)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
IsParentInactive = False
While NOT dbRS.EOF
	' added by Carlos Torres on 7/9/2020 to filter out disabled products for Tracfone users
	If (Session("pre_UserCompanyID")&"" <> "3346") Or (Session("pre_UserCompanyID")&"" = "3346" And UCase(Left(dbRS("ProductName")&" ",1)) <> "Z") Then
		If Trim(dbRS("ParentProductID")&"")<>"" Then
			If Not IsParentInactive Then
				comboList = comboList & "<option value='" & dbRS("ProductID") & "'>&nbsp;&nbsp;&nbsp;- " & Left(dbRS("ProductName"),20) & "</option>"
			End If
		Else
			comboList = comboList & "<option value='" & dbRS("ProductID") & "'>" & Left(dbRS("ProductName"),22) & "</option>"
			IsParentInactive = False
		End If
	ElseIf (Session("pre_UserCompanyID")&"" = "3346" And UCase(Left(dbRS("ProductName")&" ",1)) = "Z" And Trim(dbRS("ParentProductID")&"")="") Then
		IsParentInactive = True
	End If
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>