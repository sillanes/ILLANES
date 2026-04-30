<%
Function GetSecret()
	Dim TypeLib, fso, file, v, fpath
	
	Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
	v = Replace(Mid(CStr(TypeLib.Guid), 2, 36),"-", "")
	Set TypeLib = Nothing

	fpath = "D:\Temp\Sessions\"
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile(fpath & v, True)
	file.Close
	Set file = Nothing
	Set fso = Nothing

	GetSecret = Hide(v)

End Function

Function GetSecretOld()

	Dim Secret, v, r, b, p1, p2, p3

	Secret1 = "OyZnIV3n"
	Secret2 = "F0qftV6h"
	Randomize

	r = Int((3 - 0 + 1) * Rnd + 0)
	Select Case r
		Case 0
			b = "41"
			p1 = Right("0" & CStr(Day(Now)),2) & Right("0" & CStr(Month(Now)),2)
			p2 = CStr(Year(Now))
			p3 = Secret1 & Secret2
		Case 1
			b = "37"
			p1 = CStr(Year(Now))
			p2 = Right("0" & CStr(Month(Now)),2) & Right("0" & CStr(Day(Now)),2)
			p3 = Secret2 & Secret1
		Case 2
			b = "23"
			p1 = Right("0" & CStr(Month(Now)),2) & Right("0" & CStr(Day(Now)),2)
			p2 = CStr(Year(Now))
			p3 = Secret1 & Secret2
		Case 3
			b = "15"
			p1 = CStr(Year(Now))
			p2 = Right("0" & CStr(Day(Now)),2) & Right("0" & CStr(Month(Now)),2)
			p3 = Secret2 & Secret1
		Case Else
			b = "41"
			p1 = Right("0" & CStr(Day(Now)),2) & Right("0" & CStr(Month(Now)),2)
			p2 = CStr(Year(Now))
			p3 = Secret1 & Secret2
	End Select

	v = p1 & p3 & p2 & b
	GetSecretOld = Hide(v)

End Function

Function Hide(value)
 Dim hidden, middle, a, c, h, i, j, n, s, lowerbound, upperbound
 hidden = ""
 middle = ""
 lowerbound = (22.5 * 4 + 15 - 2)
 upperbound = (8 * 15.25)
 s = Chr(270 / 6 + 3) & Chr(8 * 13 + 21 - 5)
 Randomize
 
 a = StrReverse(value)
 For j = 1 To Len(a)
	c = Mid(a, j, 1)
	h = s & Hex(Asc(c))
	i = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
	h  = Replace(h, s, Chr(i))
	middle = middle & h
 Next
 
 a = StrReverse(middle)

 For j = 1 To Len(a)
	n = Mid(a, j, 1)
	i = Rnd
	If i = 1 Then
		n = UCase(n)
	End If
	hidden = hidden & n
 Next

 Hide = hidden
End Function

Function Reveal(value)
	Dim revealed, temp, c, d, i, j, l, n, p, s, t, u, x
	revealed = ""
	temp = ""
	t = Chr(4 * 8 + 6) & Chr(6 * 7 + 39 - 9)
	s = Chr(12 * 6 - 24) & Chr(13 * 9 + 3)
	l = (22.5 * 4 + 15 - 2)
	u = (8 * 15.25)
	
	d = StrReverse(value)
	For j = 1 To Len(d)
		n = Mid(d, j, 1)
		p = Asc(LCase(n))
		If p >= l And p <= u Then
			n = s
		End If
		temp = temp & LCase(n)
	Next
	x = Split(temp, s)
	For i = UBound(x) To LBound(x) Step -1
		If Len(CStr(x(i))) > 0 Then
			c = Chr(CLng(t & x(i)))
			revealed = revealed & c
		End If
	Next
	Reveal = revealed
End Function
%>