<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->

<%

'////////////////////////////////////////////////////
'  Լ34��: sDate = n ���n������̫�٣����n�Ĵ�Щ
'  Լ40��: nTop = 4 '20 ��ʾ��������
'  Peace ��ʾ
'////////////////////////////////////////////////////


Dim oblog,yAct
set oblog=new class_sys
oblog.autoupdate=False
oblog.start
dim js_blogurl : js_blogurl=Trim(oblog.CacheConfig(3))



yAct = Request("yAct") ' Write,View,Vjs,Index
Response.Charset="gb2312"
Response.Write "<meta http-equiv='content-type' content='text/html; charset=gb2312'>"


If yAct="Write" Then 'sub showlog()

	dim rs,sql,ars,i,s,t
	dim orders,topic,isbest
	dim postname,classid,posttime,userid
	dim usersql,isbestsql,userurl,sdatesql
	Dim sDate,nTop,jInfo,tLen
	
	'j=0&classid=all&subjectname=0&classname=0&tlen=32&n=20&sdate=7&orders=1&info=7&action=1&user=
	sDate = 7 'Int(Request("sdate"))
   	userid=0 'userid=CLng(Request("user"))
   	orders="iis" 'orders=1
   	'orders="logid" 'orders=2
   	orders="commentnum" 'orders=3
    classid=""
	nTop = 6 '20
	tLen = 35
	usersql="" '" and oblog_log.userid="&userid
    'isbestsql=" and isbest=1" 'action=2
	isbestsql="" 'action=1

	sdatesql = " and datediff('d',oblog_log.truetime,Now())<"&Int(sdate)&" and oblog_log.blog_password=0 And (ispassword='' Or ispassword is null)"

	''sdate = DateAdd("d",-1*Abs(sdate),Now())
	'sdate = GetDateCode(sdate,0)
	'sdatesql = " and truetime>'"&sdate&"' and oblog_log.blog_password=0 And (ispassword='' Or ispassword is null)"

	sql = "select top "&nTop&" author,topic,logid,classid,subjectid,truetime,iis,commentnum,logfile,oblog_log.userid,user_domain,user_domainroot from oblog_log,oblog_user where passcheck=1 and oblog_log.isdel=0 and isdraft=0 "&sdatesql&isbestsql&classid&usersql&" and oblog_user.userid=oblog_log.userid ORDER BY "&orders&" desc"
	'Response.write sql
	set rs=oblog.execute(sql)
	i=0 : s=vbcrlf&"<table width='100%' border='0' cellpadding='1' cellspacing='1'>"
	do while Not RS.Eof and i<nTop
    postname=Trim(rs(0))
    POSTTIME=rs(5)
	topic=oblog.filt_html(rs(1)) 'Replace(,"\","\\")
	if oblog.CacheConfig(5)=1 then
		userurl="http://"&rs(10)&"."&Trim(rs(11))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(9)
	end if
	if oblog.strLength(topic)>tLen then
        topic=oblog.InterceptStr(topic,tLen+3)&"..."
    end if
	
	s=s&vbcrlf&"<tr><td nowrap='nowrap' style='font-size:12px;line-height:150%'>"
	s=s&vbcrlf&"<FONT color=#800000 style=font-family:webdings>4</FONT>"
	s=s&vbcrlf&"<a href="&js_blogurl&"go.asp?logid="&rs(2)&" target=_blank title='"&POSTTIME&"'>"&topic&"</a></td>"

	's=s&"<td nowrap='nowrap' style='font-size:12px;line-height:150%'><font color=gray>("&formatdatetime(POSTTIME,1)&")</font></td></tr>"
	s=s&"<td nowrap='nowrap' style='font-size:12px;line-height:150%'><font color=gray>(����"&rs(7)&"ƪ)</font></td></tr>"
	

	RS.MoveNext
	i=i+1
	Loop
	rs.close
    set ars=nothing
	set rs=nothing
	

	  s=s&vbcrlf&"</table>"
 	  Dim st
	  Set st=Server.CreateObject("ADODB.Stream")   
  	  st.Type=2   
  	  st.Mode=3   
 	     st.Charset="gb2312"  
 	     st.Open()   
  	  st.WriteText s  
	  st.SaveToFile Server.MapPath("/blog/DataSource/showlog.htm"),2   
	  st.Close()  
 	  Set st=Nothing  
	  Response.Write s
	
	
ElseIf yAct="View" Then 'end sub

    Dim str,stm
	  set stm=server.CreateObject("adodb.stream")
	  stm.Type=2 
	  stm.Mode=3 
	  stm.Charset="gb2312"
	  stm.Open
	  stm.loadfromfile Server.MapPath("/blog/DataSource/showlog.htm")
	  str=stm.ReadText
	  stm.Close
    set stm=nothing
	Response.Write vbcrlf&str


ElseIf yAct="Vjs" Then


	Response.Write vbcrlf&"<SCRIPT language=JavaScript src='js.asp?j=0&classid=all&subjectname=0&classname=0&tlen=32&n=20&sdate=7&orders=1&info=7&action=1&user=' type=text/javascript"&"><"&"/SCRIPT>"

	'<!--
	'Peace       : sdate=n ���n������̫�٣����n�Ĵ�Щ
	'j=0         : ��ʾ������־
	'classid     : ϵͳ����id��ȫ��Ϊall
	'subjectname : 0:Ϊ������ 1:�����û�ר������ 
	'classname   : 0:Ϊ������ 1:����ϵͳ��������
	'tlen        : ���ⳤ��
	'n���� ��    : ��ʾ���ٸ�����
	'sdate ��    : ��ѯ����������־��1Ϊ����
	'orders��    : ���򷽷���1Ϊ���յ��(������־)��2Ϊ����ʱ��(�����»ظ�ʱ��),3:Ϊ����������(���ظ���־)
	'info        : 1Ϊ��ʾ����ʱ����û���2Ϊ��ʾ����ʱ�䣬3Ϊ��ʾ�����û���4Ϊ��ʾ�����û��͵������5Ϊ��ʾ�������6Ϊ��ʾ�������ں��û���7Ϊ��ʾ�������ڣ�8Ϊ��ʾ�ظ�����0Ϊ����ʾ
	'action      : 1: ��ʾ������־ 2:��ʾ������־
	'user        : �û�id���루���֣�����ʾ���û�����־(���Ե��ù���Ա��־������һ���򵥵����ŷ���ϵͳ)
	'-->


ElseIf yAct="Index" Or yAct="" Then


    Response.Write vbcrlf&"<a href='?yAct=Vjs'>Vjs</a> - <a href='?yAct=Write'>Write</a> - <a href='?yAct=View'>View</a> - <a href='?yAct=Index'>Index</a>"


End If 

%>

