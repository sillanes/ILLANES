<%
 
Dim lineData  
Dim MyArray
Dim i
i=0

'reading the file
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile("C:\Reclamos\Extractos\DetalleMovimientosHSBC.csv", 1, true)    
Do Until fs.AtEndOfStream 

	lineData = fs.ReadLine
	if i<>0 THEN
		ReDim MyArray(0)
		MyArray = Split(lineData , ",")
		Fecha= MyArray(0)
		TipoOperacion=MyArray(1)
		Comprobante=MyArray(2)
		Desc1=MyArray(3)
		Debito=MyArray(4)
		Credito=MyArray(5)
		Desc2=MyArray(6)
		Mov1=MyArray(7)
		Mov2=MyArray(8)  
				
		
		sSQL = "EXEC TMP.usp_ExtractosBancarios_Ins "
		
		FOR EACH Field IN MyArray
		
			sSQL = sSQL &"'"& Field &"',"
		NEXT
		
		response.write sSQL
		
		ReDim MyArray(0)
		
	End If
	i=i+1
Loop 

fs.close: set fs = nothing 
set o = nothing

%>