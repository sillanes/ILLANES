<%
Dim FolderSite, NameConfigFile
FolderSite = "prepaid"
NameConfigFile = "config.ini" 'If the configuration file is in the "includes" folder, put "includes/config.ini"
If IsEmpty(Session("PortalName")) Then
	Session("PortalName") = "retail"
End If

Sub IniFileLoad(ByVal FilSpc)
  Dim PhyPth
  Dim FilSys
  Dim IniFil
  Dim StrBuf
  Dim HdrBuf
  Dim StrPtr
  Dim AltBuf
  dim IniFileDictionary
  
  set IniFileDictionary = CreateObject("Scripting.Dictionary") 
  
  FilSpc = lcase(FilSpc)
  if left(FilSpc, 1) = "p" then
    'Physical path
    PhyPth = mid(FilSpc, instr(FilSpc, "=") + 1)
  else
    'Virtual path
    PhyPth = Server.MapPath(mid(FilSpc, instr(FilSpc, "=") + 1))
  end if

  set FilSys = CreateObject("Scripting.FileSystemObject")
  set IniFil = FilSys.OpenTextFile(PhyPth, 1)
  do while not IniFil.AtEndOfStream
    StrBuf = IniFil.ReadLine
    if StrBuf <> "" then
      'There is data on this line
      if left(StrBuf, 1) <> ";" then
        'It's not a comment
        if left(StrBuf, 1) = "[" then
          'It's a section header
          HdrBuf = mid(StrBuf, 2, len(StrBuf) - 2)
        else
          'It's a value
          StrPtr = instr(StrBuf, "=")
          AltBuf = lcase(HdrBuf & "|" & left(StrBuf, StrPtr - 1))
          do while IniFileDictionary.Exists(AltBuf)
            AltBuf = AltBuf & "_"
          loop
          IniFileDictionary.Add AltBuf, mid(StrBuf, StrPtr + 1)
        end if
      end if
    end if
  loop
  IniFil.Close
  set IniFil = nothing
  set FilSys = nothing
  set Session("IniFileDictionary_prepaid") = IniFileDictionary
  
End Sub

Function IniFileValue(ByVal ValSpc)
  Dim ifarray
  Dim StrPtr
  Dim StrBuf
  dim IniFileDictionary
  
  if  IsEmpty(Session("IniFileDictionary_prepaid")) then
	dim path
	path = GetPath() & "\\" & NameConfigFile

	call IniFileLoad("path=" & path)
  end if
  
  set IniFileDictionary = Session("IniFileDictionary_prepaid")
  
  StrPtr = instr(ValSpc, "|")
  ValSpc = lcase(ValSpc)
  if StrPtr = 0 then
    'They want the whole section
    StrBuf = ""
    StrPtr = len(ValSpc) + 1
    ValSpc = ValSpc + "|"
    ifarray = IniFileDictionary.Keys
    for i = 0 to IniFileDictionary.Count - 1
      if left(ifarray(i), StrPtr) = ValSpc then
        'This is from the section
        if StrBuf <> "" then
          StrBuf = StrBuf & "~"
        end if
        StrBuf = StrBuf & ifarray(i) & "=" & IniFileDictionary(ifarray(i))
      end if
    next
  else
    'They want a specific value
    StrBuf = IniFileDictionary(ValSpc)
  end if
  IniFileValue = StrBuf
End Function

Function IsToggleOn()
	' [CPI-16] Variables created for the toggle
	dim IniFileDictionary
  
	if  IsEmpty(Session("IniFileDictionary_prepaid")) then
		dim path
		path = GetPath() & "\\" & NameConfigFile

		call IniFileLoad("path=" & path)
	end if
  
	set IniFileDictionary = Session("IniFileDictionary_prepaid")
	
	dim accessType
	accessType = ""
	accessType = IniFileValue("AccessType|KeyCloak")

	If accessType = "ON" Then
		IsToggleOn = True
	Else
		IsToggleOn = False
	End If
End Function

Function IsToggleAuthOn()
  ' [CPI-16] Variables created for the toggle
	dim IniFileDictionary
  
	if  IsEmpty(Session("IniFileDictionary_prepaid")) then
		dim path
		path = GetPath() & "\\" & NameConfigFile

		call IniFileLoad("path=" & path)
	end if

	set IniFileDictionary = Session("IniFileDictionary_prepaid")

	dim accessType
	accessType = ""
	accessType = IniFileValue("AccessType|KeyCloakAuth")

	  If accessType = "ON" Then
		IsToggleAuthOn = True
	  Else
		IsToggleAuthOn = False
	  End If
End Function

Function GetKeycloakConfig(value)
	dim IniFileDictionary
  
	if  IsEmpty(Session("IniFileDictionary_prepaid")) then
		dim path
		path = GetPath() & "\\" & NameConfigFile

		call IniFileLoad("path=" & path)
	end if

	set IniFileDictionary = Session("IniFileDictionary_prepaid")
  
	GetKeycloakConfig = IniFileValue("KeycloakConfig|" & value)
End Function

Function GetPath()
    Dim currentPath, rootPath

    currentPath = Server.MapPath(".")
    rootPath = mid(currentPath, 1, InStr(currentPath, FolderSite)-1) & FolderSite
    GetPath = rootPath
End Function

function CurrentPage()
    dim s, protocol, port

    if Request.ServerVariables("HTTPS") = "on" then 
        s = "s"
    else 
        s = ""
    end if  
 
    protocol = strleft(LCase(Request.ServerVariables("SERVER_PROTOCOL")), "/") & s 

    if Request.ServerVariables("SERVER_PORT") = "80" then
        port = ""
    else
        port = ":" & Request.ServerVariables("SERVER_PORT")
    end if  

    CurrentPage = protocol & "://" & Request.ServerVariables("SERVER_NAME") & port 
end function

function StrLeft(str1,str2)
    StrLeft = Left(str1,InStr(str1,str2)-1)
end function
%>