
        <%

Function uInd_News()

Dim arr(2),ar2(2)
arr(0) = "300News"
arr(1) = "500Edu"
For i = 0 to 1
TypID = arr(i)

 sql = " SELECT TOP 10 * FROM [NewsInfo] " '
 sql =sql& " WHERE KeyID=KeyRe AND InfType LIKE '%"&TypID&"%' "
 sql =sql& " ORDER BY KeyID DESC" 'SetTop,
   rs.Open Sql,conn,1,1
     s=""
   Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyRe = rs("KeyRe")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = rs("InfType")
InfSubject = rs("InfSubject")&""
SetSub = rs("SetSub")
InfKey = rs("InfKey")
InfFrom = rs("InfFrom")
InfContent = rs("InfContent")
SetRead = rs("SetRead")
ImgName = rs("ImgName")
CType = rs_Val(conn,"SELECT TypName FROM NewsType WHERE TypLayer='"&InfType&"'","TypName")
If LEN(InfSubject)>17 Then
  InfSubject = Left(InfSubject,17)
End If
InfSubject = Show_SetSubj(InfSubject,SetHot,SetSub)

	  LayArr = Split(InfType,";")
	  LayStr = ""
	  For ix = 0 To uBound(LayArr)
	    LayStr = LayStr&"&TP"&ix&"="&LayArr(ix)
	  Next

   s=s&"[<a href='/news/news_list.asp?send="&LayStr&"' target='_blank'>"&CType&"</a>]<a href='/news/news.asp?ID="&KeyID&"' target='_blank'>"&InfSubject&"</a><br>"

   rs.MoveNext
   Loop
   rs.Close()

    set fil1 = fso.CreateTextFile(Server.MapPath("../../script/ind10/news_ind_"&TypID&".htm"),True)
    fil1.Write(s)
    fil1.Close

Next

uInd_News = s
End Function  



	  %>
          
