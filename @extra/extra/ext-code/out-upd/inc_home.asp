<%

Function uPag_Home()

PathImg = "../../script/blog/img/"
sp = ""
st = ""
   sql = " SELECT TOP 5 * FROM [HomBlog] WHERE KeyMod='Photo' AND SetSAdm='Y' "
   sql =sql& " AND LEN(HomBlog.ImgName)>12 "
   sql =sql& " ORDER BY KeyID DESC" 
   rs.Open Sql,conn,1,1
  Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
KeyRe = rs("KeyRe")
KeyMod = rs("KeyMod")
KeyUser = rs("KeyUser")
KeyFlag = rs("KeyFlag")
KeyFrom = rs("KeyFrom")
KeyTo = rs("KeyTo")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubject = Left(rs("InfSubject"),13)
InfSubject = Show_Text(InfSubject)
InfKey = rs("InfKey")
InfFrom = rs("InfFrom")
InfContent = rs("InfContent")
LogATime = rs("LogATime")
ImgName = rs("ImgName")
ImgAlign = rs("ImgAlign")
ImgTitle = rs("ImgTitle")
ImgWidth = rs("ImgWidth")
ImgHeight = rs("ImgHeight")
ImgScale = rs("ImgScale")
if LenB(InfSubj)>35 then
    InfSubj = Left(InfSubj,17)
end if
InfSubj = Show_Text(InfSubj)
if LenB(bsType)>8 then
    If LenB(TypName)=Len(TypName) Then
	  TypName = Left(TypName,4)
	Else
	  TypName = Left(TypName,8)
	End If
end if
sp = sp& "<td width='20%'><a href=/blog/photo.asp?ID="&KeyID&" target=_blank><img src='"&PathImg&ImgName&"' width='80' height='60' border='0'></a></td>"
'st = st& "&middot;<a href=/news/news.asp?TP="&InfType&"&ID="&KeyID&" target=_blank>"&InfSubj&""&"</a> <font color='#CCCCCC'>"&TypName&"</font><br>"
rs.movenext
  Loop 'next 
rs.close
	
    set fil1 = fso.CreateTextFile(Server.MapPath("../../script/ind10/home_top_img.htm"),True)
    fil1.Write(sp)
    fil1.Close
	
	
	
s = ""	
    sql = " SELECT TOP 10 * FROM [HomBlog] WHERE KeyMod='TraHom' AND SetSAdm='Y' "
	sql =sql& " ORDER BY KeyID DESC"
	
   rs.Open Sql,conn,1,1
If NOT rs.EOF Then
Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
KeyRe = rs("KeyRe")
KeyMod = rs("KeyMod")
KeyUser = rs("KeyUser")
KeyFlag = rs("KeyFlag")
KeyFrom = rs("KeyFrom")
KeyTo = rs("KeyTo")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubject = Left(rs("InfSubject"),13)
InfSubject = Show_Text(InfSubject)
InfKey = rs("InfKey")
InfFrom = rs("InfFrom")
InfContent = rs("InfContent")
LogATime = rs("LogATime")
s = s& "[<a href='/blog/suhome_list.asp?TP="&InfType&"' target='_blank'>"&InfType&"</a>]&middot;<a href='/blog/suhome.asp?ID="&KeyID&"' target='_blank'>"&InfSubject&"</a><br>"
rs.MoveNext
Loop
Else
s = s& " &nbsp; &middot; 暂无资料"
End If
rs.close
	
    set fil1 = fso.CreateTextFile(Server.MapPath("../../script/ind10/home_mid_supply.htm"),True)
    fil1.Write(s)
    fil1.Close
	
	

s = ""		
    sql = " SELECT TOP 15 * FROM [HomBlog] WHERE KeyMod='Blog' AND SetSAdm='Y' "
	sql =sql& " ORDER BY KeyID DESC"
	
   rs.Open Sql,conn,1,1
If NOT rs.EOF Then
Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
KeyRe = rs("KeyRe")
KeyMod = rs("KeyMod")
KeyUser = rs("KeyUser")
KeyFlag = rs("KeyFlag")
KeyFrom = rs("KeyFrom")
KeyTo = rs("KeyTo")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubject = Left(rs("InfSubject"),13)
InfSubject = Show_Text(InfSubject)
InfKey = rs("InfKey")
InfFrom = rs("InfFrom")
InfContent = rs("InfContent")
LogATime = rs("LogATime")
s = s& "[<a href='/blog/home.asp?CD="&KeyCode&"' target='_blank'>"&KeyCode&"</a>]&middot;<a href='/blog/arthome.asp?ID="&KeyID&"' target='_blank'>"&InfSubject&"</a><br>"
rs.MoveNext
Loop
Else
s = s& " &nbsp; &middot; 暂无资料"
End If
rs.close
	
	
    set fil1 = fso.CreateTextFile(Server.MapPath("../../script/ind10/home_mid_news.htm"),True)
    fil1.Write(s)
    fil1.Close
	
	
s = ""	
    sql = " SELECT TOP 17 * FROM [HomUser] WHERE 1=1 AND SetSAdm='Y' "
	sql =sql& " ORDER BY KeyID DESC"
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
If NOT rs.EOF Then
Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = rs("InfType")
InfName = rs("InfName")
InfNotes = rs("InfNotes")
InfFrom = rs("InfFrom")
InfCard = rs("InfCard")
LogATime = rs("LogATime")
InfMsg = InfName&"("&InfFrom&")"&""&InfType&""
InfMsg = Left(InfMsg,13)
InfMsg = Show_Text(InfMsg)
s = s& "&nbsp;<a href='/blog/home.asp?CD="&KeyCode&"' target=_blank>"&InfMsg&"</a><BR>"
rs.MoveNext
Loop
Else
s = s& " &nbsp; &middot; 暂无资料"
End If
rs.close
	
    set fil1 = fso.CreateTextFile(Server.MapPath("../../script/ind10/home_mid_home.htm"),True)
    fil1.Write(s)
    fil1.Close
	
	
uPag_Home = s
End Function  



%>