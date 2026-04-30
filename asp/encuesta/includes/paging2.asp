<%
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.cursorlocation=3 'Client=3, Server=2
dbRS.cachesize=5
If dbRS_SORT<>"" Then	dbRS.Sort = dbRS_SORT

dbRS.open sSQL, dbCon
howmanyrecs=0
Not_dbRS_EOF=False
If dbRS.State<>0 Then	'0:closed, 1:open, 2:connecting, 4:executing, 8:fetching
	If Not (dbRS.EOF or dbRS.BOF) Then
		Not_dbRS_EOF=True
		rwcnt=0
		mypage=cint("0"&Request("pg2"))
		If cint("0"&mypagesize)=0 Then mypagesize=15

		fieldCount=dbRS.Fields.Count
		recordCount=dbRS.RecordCount
		dbRS.movefirst
		dbRS.pagesize=mypagesize
		maxrecs=mypagesize
		maxpages=cint("0"&dbRS.pagecount)

		If mypage=0 Then mypage=1
		If mypage>maxpages Then mypage=maxpages
		dbRS.absolutepage=mypage
		sno=1+((mypage-1)*mypagesize)

		pgOfStr = ""
		If maxpages>1 Then	pgOfStr = "Page - " & mypage & " of " & maxpages

		pagingStr = ""
		'if mypage <> 1 then    pagingStr = pagingStr & "<a class='nav' href='First' onClick='javascript:chk_Frm2(document.FF,1);return false'><b>First</b></a>&nbsp;&nbsp;"
		if mypage > 1 then     pagingStr = pagingStr & "<a class='nav' href='Previous' onClick='javascript:chk_Frm2(document.FF," &mypage-1& ");return false'><b>Previous</b></a>&nbsp;&nbsp;"
		if maxpages > 1 then
			pgFirst = mypage - 4
			if pgFirst < 1 then	pgFirst = 1
			pgLast = pgFirst + 8
			if pgLast > maxpages then pgLast = maxpages

			pagingStr = pagingStr & "<font color='FFFFFF'>[</font>"
			pagingStr = pagingStr & "<a class='nav' href='javascript:chk_Frm2(document.FF,1)'>1...</a>&nbsp;&nbsp;"

			For i = pgFirst To pgLast
				pagingStr = pagingStr & "<a class='nav' href='javascript:chk_Frm2(document.FF," & i & ")'>" & i & "</a>&nbsp;&nbsp;"
			Next

			pagingStr = pagingStr & "&nbsp;<a class='nav' href='javascript:chk_Frm2(document.FF," &maxpages& ")'>..." &maxpages& "</a>"
			pagingStr = pagingStr & "<font color='FFFFFF'>]</font>&nbsp;"
		end if
		if mypage < maxpages then   pagingStr = pagingStr & "&nbsp;<a class='nav' href='Next' onClick='javascript:chk_Frm2(document.FF," &mypage+1& ");return false'><b>Next</b></a>&nbsp;&nbsp;"
		'if mypage <> maxpages then  pagingStr = pagingStr & "<a class='nav' href='Last' onClick='javascript:chk_Frm2(document.FF," &maxpages& ");return false'><b>Last</b></a>&nbsp;"
	End If
End If
%>