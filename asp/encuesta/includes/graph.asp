<%
IF db_getRSdata Then
	If sSQL<>"" Then
		Set RS=Server.CreateObject("ADODB.Recordset")
		RS.cursorlocation=3
		RS.cachesize=5
		RS.open sSQL,dbCon
		If NOT RS.EOF Then	
			img_draw = true
			array_Range1=RS.RecordCount
		End If
		array_Range=array_Range1
	End If
	
	If sSQL2<>"" Then
		Set RS2=Server.CreateObject("ADODB.Recordset")
		RS2.cursorlocation=3
		RS2.cachesize=5
		RS2.open sSQL2,dbCon
		If NOT RS2.EOF Then	
			img_draw2 = true
			array_Range2=RS2.RecordCount
		End If
		If array_Range1<array_Range2 Then array_Range=array_Range2
	End If
	
	IF img_draw Or img_draw2 THEN
		rem ***************** Recordset Display functions *********************
		If x_DisplayFunction="" Then x_DisplayFunction = "FormatDateTime(RS(0),VBShortTime)"
		If y_DisplayFunction="" Then y_DisplayFunction = "CLng(RS(1))"
		If x_DisplayFunction2="" Then x_DisplayFunction2 = "FormatDateTime(RS2(0),VBShortTime)"
		If y_DisplayFunction2="" Then y_DisplayFunction2 = "CLng(RS2(1))"

		rem ***************** Declare X Array *********************************
			ReDim myArrx(array_Range)
		
		rem ***************** Store X and Y values into an Array1 *************
		If img_draw Then
			ReDim myArr(array_Range1)
		
			i=0
			yMax1=0
			While NOT RS.EOF
				myArrx(i)=Eval(x_DisplayFunction)
				temp=Eval(y_DisplayFunction)
				myArr(i)=temp
				if yMax1<temp then yMax1=temp
				i=i+1
				RS.MoveNext
			Wend
			RS.Close
			SET RS=Nothing
			xy_points=i
		End If
		
		rem ***************** Store X and Y values into an Array2 *************
		If img_draw2 Then
			ReDim myArr2(array_Range2)
		
			i=0
			yMax2=0
			While NOT RS2.EOF
				myArrx(i)=Eval(x_DisplayFunction2)
				temp=Eval(y_DisplayFunction2)
				myArr2(i)=temp
				if yMax2<temp then yMax2=temp
				i=i+1
				RS2.MoveNext
			Wend
			RS2.Close
			SET RS2=Nothing
			xy_points2=i
		End If
	End If
End If

IF img_draw Or img_draw2 THEN
	rem ***************** Set min and max points **************************
	If yMax2<yMax1 Then 
		yMax=yMax1
	Else
		yMax=yMax2
	End If
	If img_draw3 Then
		If img_draw AND img_draw2 AND array_Range1=array_Range2 Then
			yMax = yMax1 + yMax2
		Else
			img_draw3 = false
		End If
	End If
	mn=myArrx(0)
	mx=myArrx(array_Range-1)
	mxf=(Fix(yMax/y_Multiplier)+1) * y_Multiplier
	
	rem **************** Image Creation ***********************************
	Dim Image
	Set Image=Server.CreateObject("AspImage.Image")
	Image.AntiAliasText=True
	Image.PadSize=0
	Image.MaxX=img_width
	Image.MaxY=img_height
	
	rem ***************	Adjust Background Color ***************************
	If img_Bg2Way Then
		Image.GradientTwoWay img_Bg2Begin, img_Bg2End, img_Bg2Direction, img_Bg2InOut
	ElseIf img_Bg1Way Then
		Image.GradientTwoWay img_Bg1Begin, img_Bg1End, img_Bg1Direction
	End If
	
	rem ************************** Text Font ******************************
	Image.FontName=y_fontFace
	Image.FontSize=y_fontSize
	Image.FontColor=y_fontColor
	
	rem *********************** Put Y Values ******************************
	pls = (img_height-bottom_space-top_space) / (y_points)
	plv = mxf/y_points
	Image.TextOut "0", y_left_space, (img_height-bottom_space)-5, false
	X0 = left_space+1
	X1 = img_width-right_space-1
	FOR i=1 to y_points
		Y0 = img_height - bottom_space - CLng(i*pls)
		Image.TextOut CLng(plv*i), y_left_space, Y0-5, false
		If i<y_points Then
			Image.PenColor = RGB(&h93, &h93, &h93)
			Image.X = X0
			Image.Y = Y0+2
			Image.LineTo X1, Y0+2
			Image.PenColor = vbBlack
		End If
	NEXT
	
	rem ************************** Text Font ******************************
	Image.FontName=x_fontFace
	Image.FontSize=x_fontSize
	Image.FontColor=x_fontColor
	
	rem *********************** Put X Values ******************************
	Image.TextAngle = 90
	x_pix = (img_width-right_space-left_space) / (x_points+1)
	plv   = (array_Range-1) / (x_points+1)
	
	X0 = left_space-5
	Y0 = x_bottom_space
	Image.TextOut myArrx(0), X0-3, Y0, false
	FOR i=1 to x_points
		Image.TextOut myArrx(CLng(plv*i)), CLng(X0 + (i*x_pix))-3, Y0, false
	NEXT
	Image.TextOut myArrx(array_Range-1), CLng(X0 + ((x_points+1)*x_pix))-3, Y0, false
	
	rem *********************** Draw Graphic ******************************
	x_pix = (img_width-left_space-right_space)/(array_Range-1)
	k  = (img_height-bottom_space-top_space)/mxf
	orj= img_height - bottom_space
	
	If img_fill Then
		Image.PenWidth = 1
		Image.PenColor = img_fillcolor
		X0 = left_space
		Y0 = (img_height-bottom_space)-2
		FOR i=1 to array_Range1-1
			X0 = X0 + x_pix 
			Y1 = orj - CLng(myArr(i) * k)
			Image.X = X0
			Image.Y = Y0
			Image.LineTo X0, Y1
		NEXT
	End If
	
	If img_draw Then
		Image.PenColor = img_PenColor
		Image.PenWidth = img_PenWidth
		X0 = left_space
		YO = orj - CLng(myArr(0) * k)
		Image.X = left_space
		Image.Y = orj - CLng(myArr(0) * k)
		FOR i=1 to array_Range1-1
			X0 = X0 + x_pix 
			Y0 = orj - CLng(myArr(i) * k)
			Image.LineTo X0, Y0
		NEXT
	End If
	
	If img_draw2 Then
		Image.PenColor = img_PenColor2
		Image.PenWidth = img_PenWidth2
		X0 = left_space
		YO = orj - CLng(myArr2(0) * k)
		Image.X = left_space
		Image.Y = orj - CLng(myArr2(0) * k)
		FOR i=1 to array_Range2-1
			X0 = X0 + x_pix 
			Y0 = orj - CLng(myArr2(i) * k)
			Image.LineTo X0, Y0
		NEXT
	End If
	
	If img_draw3 Then
		Image.PenColor = img_PenColor3
		Image.PenWidth = img_PenWidth3
		X0 = left_space
		YO = orj - CLng((myArr(0)+myArr2(0)) * k)
		Image.X = left_space
		Image.Y = orj - CLng((myArr(0)+myArr2(0)) * k)
		FOR i=1 to array_Range-1
			X0 = X0 + x_pix 
			Y0 = orj - CLng((myArr(i)+myArr2(i)) * k)
			Image.LineTo X0, Y0
		NEXT
	End If
	
	rem *************** Draw Rectangular **********************************
	Image.BackgroundColor=img_BorderColor
	Image.FrameRect left_space, top_space, img_width-right_space, img_height-bottom_space
	Image.FrameRect left_space+1, top_space+1, img_width-right_space+1, img_height-bottom_space+1
	
	rem *********************** Write Image Titles ************************
	If img_Title<>"" Then
		Image.TextAngle=0
		Image.FontName=img_Title_fontFace
		Image.Bold=img_Title_Bold
		Image.FontSize=img_Title_fontSize
		Image.TextOut img_Title, img_Title_x, img_Title_y, false
	End If
	If img_yTitle<>"" Then
		Image.TextAngle=90
		Image.FontName=img_yTitle_fontFace
		Image.Bold=img_yTitle_Bold
		Image.FontSize=img_yTitle_fontSize
		Image.TextOut img_yTitle, img_yTitle_x, img_yTitle_y, false
	End If
	If img_xTitle<>"" Then
		Image.TextAngle=0
		Image.FontName=img_xTitle_fontFace
		Image.Bold=img_xTitle_Bold
		Image.FontSize=img_xTitle_fontSize
		Image.TextOut img_xTitle, img_xTitle_x, img_xTitle_y, false
	End If
	
	rem *********************** Save & Show Image *************************
	Image.ImageFormat=img_imageFormat
	Image.FileName=img_filePath & img_fileName
	Image.SaveImage
	Set Image = Nothing
	If img_show Then Response.Write "<img src='" & img_fileURL & img_fileName & "'>"
END IF
%>