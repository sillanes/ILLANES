<%

' 1️⃣ Separar ruta + archivo
Dim fullPath
fullPath=Session("FileUploaded")
ruta = Left(fullPath, InStrRev(fullPath, "\") - 1)
archivo = Mid(fullPath, InStrRev(fullPath, "\") + 1)

' 2️⃣ Separar nombre + extensión
nombre = Left(archivo, InStrRev(archivo, ".") - 1)
extension = Mid(archivo, InStrRev(archivo, ".") + 1)

sSQL = "EXEC [cobranza].[Transportista_Comprobante_Ins] " & hdrid & ", " & clienteid & ",'" & nombre&"."&extension  & "'" & ",'" & ruta  & "'" & ",'" & extension & "'" & ",'" & Session("NombreTransportista") & "'" 
' Response.write(sSQL)
conn.Execute(sSQL)
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If

%>

