<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%
'PeaceEdit//////////////////////
Dim PagePrev,PageThis
PagePrev = LCase(Request.Servervariables("HTTP_REFERER"))
PageThis = LCase(Request.ServerVariables("SERVER_NAME")) 
PageThis = Replace(PageThis,"http://","")
PageThis = Replace(PageThis,"www.","")
If not inStr(PagePrev,PageThis)>0 Then
  'Response.Redirect "/blog/"
End If	
'PeaceEnd////////////////////////
Dim oblog
set oblog=new class_sys
oblog.autoupdate=False
oblog.start
dim js_blogurl,n
js_blogurl=Trim(oblog.CacheConfig(3))
n=CInt(Request("n"))
if n=0 then n=1
select case CInt(Request("j"))
	case 1
		call tongji()
	case 2
		call topuser()
	case 3
		call adduser()
	case 4
		call listclass()
	case 5
		call showusertype()
	case 6
		call listbestblog()
	case 7
		call showlogin()
	case 8
		call showplace()
	case 9
		call showphoto()
	case 10
		call showblogstars()
	Case 11
		Call show_hotblog()
	Case 12
		Call show_teams()
	Case 13
		Call show_posts()
	Case 14
		Call show_hottag()
	Case 15
		Call showplac2()
	Case 16
		Call show4Pic()
	case 0
		call showlog()
end select

sub tongji()
	dim rs,logcount,commentcount,messagecount,usercount
	dim today_log,yesterday_log
	set rs=oblog.execute("select log_count,comment_count,message_count,user_count,log_count_Yesterday from oblog_setup")
	logcount=rs(0)
	commentcount=rs(1)
	messagecount=rs(2)
	usercount=rs(3)
	yesterday_log = rs(4)
	If Is_Sqldata = 0 Then
		Set rs = oblog.execute("select COUNT(logid) FROM oblog_log WHERE DATEDIFF('d',truetime,Now)=0 AND isdel=0 ")
	Else
		Set rs = oblog.execute("select COUNT(logid) FROM oblog_log WHERE truetime>=CONVERT(CHAR(10),GETDATE(),120) AND truetime < CONVERT(CHAR(10),GETDATE()+1,120) AND isdel=0 ")
	End if
	today_log=rs(0)
	%>
 document.write('◎- 博客总数 <font color=green><%=usercount%></font><br> ◎- 日志总数 <font color=green><%=logcount%></font><br> ◎- 评论总数 <font color=green><%=commentcount%></font><br> ◎- 留言总数 <font color=green><%=messagecount%></font>');
 document.write('<br> ◎- 今天日志 <font color=red><%=Today_log%></font><br> ◎- 昨天日志 <font color=green><%=OB_IIF(yesterday_log,0)%></font></font>')
<%
	set rs=nothing
end sub

sub topuser()
	dim i,blogname,rs,userurl,order,ordersql
	order=CLng(Request("order"))
	i=0
	if order=0 then
		ordersql="log_count"
	ElseIf order=1 Then
		ordersql="user_siterefu_num"
	ElseIf order=2 Then
		ordersql="scores"
	end if
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot from [oblog_user] order by "&ordersql&" desc")
	do while Not RS.Eof and n>i
	if Trim(rs(2))<>"" then
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	else
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	end if
	if oblog.CacheConfig(5)=1 then
		userurl="http://"&rs(4)&"."&Trim(rs(5))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(3)
	end if
	Response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"&rs(0)&"的blog页>');"
    Response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
	Response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub


sub adduser()
	dim i,blogname,rs,userurl
	i=0
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot from [oblog_user] order by userid desc")
	do while Not RS.Eof and n>i
	if Trim(rs(2))<>"" then
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	else
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	end if
	if oblog.CacheConfig(5)=1 then
		userurl="http://"&rs(4)&"."&Trim(rs(5))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(3)
	end if
	Response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"&rs(0)&"的blog页>');"
    Response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
	Response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub

sub listbestblog()
	dim i,blogname,rs,userurl
	i=0
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot from [oblog_user] where user_isbest=1 order by log_count desc")
	do while Not RS.Eof and n>i
	if Trim(rs(2))<>"" then
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	else
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	end if
	if oblog.CacheConfig(5)=1 then
		userurl="http://"&rs(4)&"."&Trim(rs(5))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(3)
	end if
	Response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"&rs(0)&"的blog页>');"
    Response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
	Response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub

sub showlogin()
	Response.Write("function chkdiv(divid){var chkid=document.getElementById(divid);if(chkid != null){return true; }else {return false; }}"&VbCrLf)
	Response.write "document.write('<div id=""ob_login""></div><script src="&js_blogurl&"inc/main.js></script><script src="&js_blogurl&"login.asp?action=showjs&injs=1></script>');"
end sub

sub showplace()
	Response.write oblog.htm2js (oblog.setup(5,0),True)
end sub

sub showplac2()
	Dim rsPeace
	set rsPeace=oblog.execute("select site_friends from oblog_setup")
	Response.write oblog.htm2js(rsPeace(0),True)
	set rsPeace=nothing
end sub

sub showusertype()
	dim rs
	set rs=oblog.execute("select id,classname from [oblog_userclass] order by RootID,OrderID")
	do while Not RS.Eof
	Response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><a href="&js_blogurl&"listblogger.asp?usertype="& rs(0) &" target=_blank title="&rs(1)&"的博客列表>');"
    Response.write "document.write('"&rs(1)&"</a> &nbsp; ');"
	Response.write "document.write('');" '<br> <span style=""font-size:9pt;line-height: 15pt""> </span>
	rs.MoveNext
	Loop
	set rs=nothing
end sub

sub listclass()
	dim rs
	Dim t
	t=CLng(Request("t"))
	set rs=oblog.execute("select id,classname from [oblog_logclass] WHERE idtype= "&t&" order by RootID,OrderID")
	do while Not RS.Eof
	Response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><a href="&js_blogurl&"list.asp?classid="& rs(0) &" target=_blank title="&rs(1)&"的日志列表>');"
    Response.write "document.write('"&rs(1)&"</a> &nbsp; ');"
	Response.write "document.write('');" '</span><br> <FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt"">
	rs.MoveNext
	Loop
	set rs=nothing
end sub

sub showlog()
	dim rs,sql,ars,i
	dim orders,topic,isbest
	dim postname,classid,posttime,userid
	dim usersql,isbestsql,userurl,sdatesql
	Dim sDate
	sDate = Int(Request("sdate"))
	if Request("user")<>"" then
   		userid=CLng(Request("user"))
	else
		userid=0
	end if
	if Trim(Request("orders"))=1 then
		orders="iis"
	elseif Trim(Request("orders"))=2 then
		orders="logid"
	elseif Trim(Request("orders"))=3 then
		orders="commentnum"
	else
		Response.Write("错误的参数")
		Response.End()
	end if
	if Trim(Request("classid"))="all" then
            classid=""
	else
		if isnumeric(Request("classid")) then
			classid=" and classid="&cint(Trim(Request("classid")))&""
		else
			Response.Write("错误的参数")
			Response.End()
		end if
    end if
	if userid>0 then
		usersql=" and oblog_log.userid="&userid
	else
		usersql=""
	end if
	if not isnumeric(Request("sdate")) then
		Response.Write("错误的参数")
		Response.End()
	end if
	if not isnumeric(Request("n")) then
		Response.Write("错误的参数")
		Response.End()
	elseif cint(Request("n"))>100 then
		Response.Write("不能调用大于100条数据")
		Response.End()
	end if

	if cint(Request("action"))=2 then
		isbestsql=" and isbest=1"
	else
		isbestsql=""
	end If
	If Is_Sqldata = 0 Then
		sdatesql = sdatesql&" and datediff('d',oblog_log.truetime,Now())<"&Int(sdate)&" and oblog_log.blog_password=0 And (ispassword='' Or ispassword is null)"
	Else
		sdate = DateAdd("d",-1*Abs(sdate),Now())
		sdate = GetDateCode(sdate,0)
		sdatesql = sdatesql&" and truetime>'"&sdate&"' and oblog_log.blog_password=0 And (ispassword='' Or ispassword is null)"
	End IF
	set rs=oblog.execute("select top "&n&" author,topic,logid,classid,subjectid,truetime,iis,commentnum,logfile,oblog_log.userid,user_domain,user_domainroot from oblog_log,oblog_user where passcheck=1 and oblog_log.isdel=0 and isdraft=0 "&sdatesql&isbestsql&classid&usersql&" and oblog_user.userid=oblog_log.userid ORDER BY "&orders&" desc")
	i=0
	do while Not RS.Eof and i<cint(Request("n"))
    postname=Trim(rs(0))
    POSTTIME=rs(5)
	topic=oblog.filt_html(Replace(Replace(Replace(Replace(rs(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	if oblog.CacheConfig(5)=1 then
		userurl="http://"&rs(10)&"."&Trim(rs(11))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(9)
	end if
	if oblog.strLength(topic)>Cint(Request("tlen")) then
        topic=oblog.InterceptStr(topic,Request("tlen")+3)&"..."
    end if
	Response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt"">');"
	if Request("classname")=1 then
	set ars=oblog.execute("select classname from oblog_logclass where id="&rs(3))
		if not ars.eof then
			Response.write "document.write('<a href="&js_blogurl&"list.asp?classid="&rs(3)&" target=_blank>〖"&oblog.filt_html(ars(0))&"〗</a>');"
		end if
	end if

	if Request("subjectname")=1 then
	set ars=oblog.execute("select subjectname from oblog_subject where subjectid="&rs(4))
		if not ars.eof then
			Response.write "document.write('<a href="&js_blogurl&"blog.asp?name="&rs(0)&"&subjectid="&rs(4)&" target=_blank>["&oblog.filt_html(ars(0))&"]</a>');"
		end if
	end if
    Response.write "document.write('<a href="&js_blogurl&"go.asp?logid="&rs(2)&" title="&topic&" target=_blank>');"
    Response.write "document.write('"&topic&"');"
	Response.write "document.write('</a>');"

	select case cint(Request("info"))
	case 0
	case 1
	Response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,0)&"</font>)');"
	case 2
	Response.write "document.write('<font color=gray>("&POSTTIME&")</font>');"
	case 3
	Response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>)');"
	case 4
	Response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&rs(6)&"</font>)');"
	case 5
	Response.write "document.write('(<font color=green>"&rs(6)&"</font>)');"
	case 6
	Response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
	case 7
	Response.write "document.write('<font color=gray>("&formatdatetime(POSTTIME,1)&")</font>');"
	case 8
	Response.write "document.write('(<font color=green>"&rs(7)&"</font>)');"
	case else

	end select

	Response.write "document.write('</span><br>');"
	RS.MoveNext
	i=i+1
	Loop
	rs.close
    set ars=nothing
	set rs=nothing
end sub

sub showphoto()
	dim rs,n,i,w,h,show_newphoto,imgsrc,fso
	Set fso = Server.CreateObject(oblog.CacheCompont(1))
	n=CLng(Request("n"))
	i=CLng(Request("i"))
	w=CLng(Request("w"))
	h=CLng(Request("h"))
	if i=1 then i="<br />" else i=""
	set rs=oblog.execute("select top "&CLng(n)&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1  order by fileid desc")
	while not rs.eof
		imgsrc=rs(0)
		imgsrc=Replace(imgsrc,right(imgsrc,3),"jpg")
		imgsrc=Replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
		if  not fso.FileExists(Server.MapPath(imgsrc)) then
			imgsrc=rs(0)
		end if
		show_newphoto=show_newphoto&"<a href='"&js_blogurl&rs("user_dir")&"/"&rs("user_folder")&"/cmd."&f_ext&"?uid="&rs("userid")&"&do=album' target='_blank'><img src="""&js_blogurl&imgsrc&""" width="""&w&""" height="""&h&""" hspace=""6"" border=""0"" vspace=""6"" alt='"&oblog.filt_html(rs(1))&"' /></a>"&i
		rs.movenext
	wend
	Response.write "document.write('"&Replace(Replace(Replace(Replace(show_newphoto,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');"
	set rs=nothing
end sub
sub showblogstars()
	dim rs,n,i,w,h,show_blogstars
	n=CLng(Request("n"))
	i=CLng(Request("i"))
	w=CLng(Request("w"))
	h=CLng(Request("h"))
	show_blogstars=show_blogstar2(n,i,w,h)
	Response.write "document.write('"&Replace(Replace(Replace(Replace(show_blogstars,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');"
	set rs=nothing
end sub

Public Function show_blogstar2(iNumber,iPerline,iWidth,iHeight)
	Dim rs,iCount,sLine : sLine=""
	If Not IsNumeric(iNumber) Then
		iNumber=1
	Else
		iNumber=CLng(iNumber)
	End If
	If iNumber=0 Then iNumber=1
	set rs=oblog.execute("select top " & iNumber & " * from oblog_blogstar where ispass=1 order by id desc")
	if not rs.eof then
			iCount=1
			Do While Not rs.Eof
				sLine = sLine & "<a href='"&rs("userurl")&"' target='_blank'>"&oblog.filt_html(rs("blogname"))&"</a><BR>" & VBCRLF
				iCount = iCount+1
				rs.MoveNext
			Loop
		show_blogstar2 = sLine
	Else
		show_blogstar2=" "
	End If
	rs.Close
	set rs=nothing
End Function
'最受欢迎的用户,计算方法
'user_siterefu_num+comment_count*1.5+message_count*1.5+sub_num*3
'访问数+回复数*1.5+留言数*1.5+被订阅数*3
't 是否显示用户头像
Sub show_hotblog()
	Dim t,GetHotUsers
	t=CLng (Request("t"))
	Dim rs, userurl,userico,i
    set rs=oblog.execute("select top "&n&" username,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder,user_icon1 from [oblog_user] where lockuser=0 and isdel=0 order by (user_siterefu_num+comment_count*1.5+message_count*1.5+sub_num*3) desc")
    GetHotUsers="<ul>"
    While Not rs.EOF
        If oblog.cacheConfig(5) = 1 Then
            userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
			js_blogurl = ""
        Else
            userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
        End If
        If t=1 Then userico="<img src=""" & OB_IIF(rs(8),""&js_blogurl&"images/ico_default.gif") & """ width=""48"" height=""48"" border='0' /><br />"
        GetHotUsers=GetHotUsers&"<li><a href="&js_blogurl&userurl&" target=""_blank"">"&userico& rs(2)&"</a></li>" & vbcrlf
        rs.MoveNext
    Wend
    GetHotUsers=GetHotUsers&"</ul>"
	GetHotUsers=oblog.htm2js(GetHotUsers,True)
    Set rs = Nothing
	Response.Write GetHotUsers
End Sub

'x:1- 最新创建/2-最活跃群组(贴数最多)/3-规模大(人数最多) / 4-推荐群组
'n: 数目
'l: 题目显示长度
'y: 是否显示图标
'w:	图标宽度，不写则默认50
'h: 图标高度，不写则默认50
Sub show_teams()
	Dim x,l,y,w,h
	x=CLng(Request("x"))
	w=(Request("w"))
	h=(Request("h"))
	l=CLng(Request("l"))
	y=Request("y")
	Dim rs,Sql,sRet,sIco
	Sql="select top " & n & " teamid,t_name,t_ico,icount0,(icount1+icount2) From oblog_team Where istate=3 and isdel=0  "
	select Case x
		Case 1
			Sql= Sql & " Order By teamid Desc"
		Case 2
			Sql= Sql & " Order By (icount1+icount2) Desc"
		Case 3
			Sql= Sql & " Order By icount0 Desc"
		Case 4
			Sql= Sql & " and isbest=1"
	End select
	Set rs=oblog.Execute(Sql)
	sRet="<div><ul>"
	Do While Not rs.Eof
		sRet=sRet & "<li>"
		If y=1 Then
			If w="" Then w=50:h=50
			sIco=LCase(Ob_IIF(rs(2),"images/ico_default.gif"))
			If Left(sico,7)<>"http://" Then sico=blogdir & sico
			sRet=sRet & "<img src=""" & sico & """ width=""" & w &""" height=""" & h &"""/><br />"
		End if
		sRet=sRet & "<a href="""&js_blogurl&"group.asp?gid=" & rs(0) & """ target=""_blank"">" & Left(oblog.filt_html((rs(1))),l) & "</a>(" & rs(3) & "/" & rs(4) & ")"
		sRet=sRet & "</li>" & Vbcrlf
		rs.movenext
	Loop
	Set rs=Nothing
	sRet=sRet & "</ul></div>"
	Response.WRITE oblog.htm2js (sRet,True)
End Sub

'获取群组文章
'teamid: 0 所有群组;如果是选择多个群组,则把群组ID用|分隔开,如1|2|8
'postnum: 帖子数目
'l:帖子主题显示字数
'u:是否显示用户名 0/1
't:是否显示发帖时间 0/1
Sub show_posts()
	Dim teamid,postnum,l,u,t
	teamid=Request("tid")
	postnum=n
	l=CInt(Request("l"))
	u=CInt(Request("u"))
	t=CInt(Request("t"))
	Dim rs,sql,sRet,sAddon
	Sql="select Top " & postnum & " teamid,postid,topic,addtime,author,userid From oblog_teampost Where idepth=0 and isdel=0 "
	If teamid<>"" And teamid<>"0" Then
		teamid=Replace(teamid,"|",",")
		teamid  = FilterIDs(teamid)
		If teamid = "" Then Exit Sub
		Sql=Sql & " And teamid In (" & teamid & ") "
	End If
	Sql=Sql & " Order by postid Desc"
	Set rs=oblog.Execute(Sql)
	sRet="<ul>"
	Do While Not rs.Eof
		sAddon=""
		sRet=sRet & "<li><a href="""&js_blogurl&"group.asp?gid=" & rs(0) & "&pid=" & rs(1) & """ target=""_blank"">" & oblog.Filt_html(Left(rs(2),l)) & "</a>"
		If u=1 Then sAddon=rs(4)
		if t=1 Then
			If sAddon<>"" Then sAddon=sAddon & ","
			sAddon=sAddon & rs(3)
		End If
		If sAddon<>"" Then sAddon="(" & sAddon & ")"
		sRet=sRet & sAddon & "</li>"
		rs.Movenext
	Loop
	Set rs = Nothing
	sRet=sRet & "</ul>"
	Response.write oblog.htm2js (sRet,True)
End Sub

'获取标签
's 表现形式 1-列表形式,2-云图形式
'n 标签数目
'x 排序方式 0 自然序/1频度最高
'y 每行显示数目
Sub show_hottag()
	Dim s,x,y
	s=CInt(Request("s"))
	x=CInt(Request("x"))
	y=CInt(Request("y"))
	Dim sContent,sSql,rst,iFont,iFontSize,i,iFontFamily
	Dim sSplit
	sSplit="&nbsp;&nbsp;"
 	sSql="select top "& n & " * From oblog_Tags Where iNum>0 "
 	If x=1 Then sSql= sSql & " Order By iNum Desc"
 	Set rst=oblog.Execute(sSql)
 	If rst.Eof Then
 		sContent=""
	Else
		Do While Not rst.Eof
			If s=1 Then
				sContent= sContent & "<font class=tag0><a href=""tags.asp?tagid=""" & rst("tagID") &""">" & rst("Name")& "(" & rst("iNum") &  ")</a></font>" & sSPlit
			Else
				iFont=rst("iNum") Mod 100
				If iFont=0 Then iFontSize=10
				If iFont>-1 And iFont<20 Then iFontSize=10 + iFont
				if iFontSize>18 and iFontSize<23 then iFontSize=20
				if iFontSize>23 and iFontSize<28 then iFontSize=25
				if iFontSize>28 then iFontSize=30
				if iFontSize >18 then iFontFamily="黑体,"
				sContent= sContent & "<a href="""&js_blogurl&"tags.asp?tagid=" & rst("tagID") & """ title="""& rst("Name") &"""><font style=""font-size:"& iFontSize &"px;line-height:26px;font-family:"&iFontFamily&"Arial, Helvetica"">" & Left(rst("Name"),10)& "</font></a>" & sSPlit
			End If
			i=i+1
			If i Mod y = 0 Then
				sContent = sContent &  "<BR/>"
			End If
			rst.Movenext
		Loop
	End If
	rst.Close
	Set rst=Nothing
	Response.write oblog.htm2js (sContent,True)
End Sub
sub show4Pic()
	dim rs,show_newphoto,imgsrc,sPics,sLnks,sTxts,sStr3
	sPics="(-^-)"&"var pics=(-`-)"
	sLnks="(-^-)"&"var mylinks=(-`-)"
	sTxts="(-^-)"&"var texts=(-`-)"
	set rs=oblog.execute("select top 4 file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1  order by fileid desc")
	while not rs.eof
		imgsrc=rs(0)
		imgsrc=Replace(imgsrc,right(imgsrc,3),"jpg")
		imgsrc=Replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
		sPics=sPics&imgsrc&"|"
		sLnks=sLnks&js_blogurl&rs("user_dir")&"/"&rs("user_folder")&"/cmd."&f_ext&"?uid="&rs("userid")&"&do=album|"
		sTxts=sTxts&oblog.filt_html(rs(1))&"|"
		rs.movenext
	wend
	set rs=nothing
		sPics=sPics&"(-`-)"
		sLnks=sLnks&"(-`-)"
		sTxts=sTxts&"(-`-)"
		sStr3=sPics&sLnks&sTxts
		sStr3=Replace(sStr3,"|(-`-)","(-`-)")
		sStr3=Replace(sStr3,"\","")
		sStr3=Replace(sStr3,"'"," ")
		sStr3=Replace(sStr3,""""," ")
		sStr3=Replace(sStr3,vbcr," ")
		sStr3=Replace(sStr3,vblf," ")
		sStr3=Replace(sStr3,"(-`-)","""")
		sStr3=Replace(sStr3,"(-^-)",VBCRLF)
		Response.write sStr3
end sub

%>