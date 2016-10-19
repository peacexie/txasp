<!--#include file="binc/_config.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=bbsName%></title>
<link rel="stylesheet" type="text/css" href="bimg/style.css">
</head>
<body>
<!--#include file="_itop.asp"-->
<div align="center" style="width:980px; height:auto; margin:auto; background-color:#F2F6FB">
  <table width="980" border="0" align="center" cellpadding="1" cellspacing="8">
    <tr>
      <td align="left" bgcolor="#FFFFFF">&nbsp;<a href="../">首页</a> &gt;&gt; <%=bbsName%> &gt;&gt; 论坛首页 </td>
      <td width="60%" align="left" bgcolor="#FFFFFF"><!--#include file="../upfile/sys/para/bbsnote.htm"--></td>
    </tr>
  </table>
  <!-- Main AAAA /////////////////////////////// -->
  <%
TP = Request("TypID")
sql = " SELECT * FROM [WebTyps] WHERE TypMod='"&ModID&"' AND TypDeep=1 ORDER BY TypID " 
rs.Open Sql,conn,1,1
sTypID="" : sTName=""
Do While NOT rs.EOF
  sTypID = sTypID&rs("TypID")&";"
  sTName = sTName&rs("TypName")&";"
  rs.Movenext
Loop
rs.close()
aTypID=Split(sTypID,";")
aTName=Split(sTName,";")
For j=0 To uBound(aTypID)-1
%>
  <table width="960" border="0" align="center" cellpadding="1" cellspacing="1" class="sysTabC1">
    <tr>
      <td align="left" class="sysTopBG fntFFF"><strong>&nbsp;<a href="?TypID=<%=aTypID(j)%>" class="LnkWhite"><%=aTName(j)%>&nbsp;<img src="bimg/collapsed_norm.gif" width="17" height="17" align="absmiddle" /></a></strong></td>
    </tr>
    <%If 1<>2 Then 'TP=aTypID(j) Or (TP="" And j=0)%>
    <tr>
      <td height="60" align="left" valign="top" class="sysTabGB"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="sysTabC2">
          <tr>
            <th width="5%" align="center" bgcolor="#FFFFFF">&nbsp;</th>
            <th width="45%" align="center" bgcolor="#FFFFFF">版块名称 </th>
            <!--<th width="5%" align="center" bgcolor="#FFFFFF">&nbsp;</th>-->
            <th align="center" bgcolor="#FFFFFF">主题/回复</th>
            <th width="25%" align="center" bgcolor="#FFFFFF">最后发表</th>
            <th width="8%" align="center" nowrap="nowrap" bgcolor="#FFFFFF">&nbsp;&nbsp;版 主&nbsp;&nbsp;</th>
          </tr>
          <%
  sql = " SELECT * FROM [WebTyps] WHERE TypLayer LIKE '"&aTypID(j)&"%' "
  sql =sql& " AND TypDeep=2  ORDER BY TypID " 
  rs.Open Sql,conn,1,1
  Do While NOT rs.EOF
   TypID = rs("TypID")
   TypLayer = rs("TypLayer")
   TypName = rs("TypName")
   TypNam2 = rs("TypNam2")&""
   ImgName = rs("ImgName")
   TypResume = rs("TypResume")
   
mReply = rs_Count(conn," BBSInfo WHERE KeyMod='"&ModID&"' AND InfType='"&TypLayer&"' AND LEN(KeyRE)>12 ")
mPub = rs_Count(conn," BBSInfo WHERE KeyMod='"&ModID&"' AND InfType='"&TypLayer&"' ")-mReply
nDay = rs_Count(conn," BBSInfo WHERE InfType='"&TypLayer&"' AND LogATime>"&cfgTimeC&""&Date()&""&cfgTimeC&" AND LEN(KeyRE)<12")
If CStr(nDay)="0" Then
  nImg = "Board1"
Else
  nImg = "Board2"
End If

   Set rs2=Server.Createobject("Adodb.Recordset")
   rs2.Open "SELECT TOP 1 * FROM BBSInfo WHERE InfType='"&TypLayer&"' AND LEN(KeyRE)<12 AND SetShow='Y' ORDER BY KeyID DESC",conn,1,1
if NOT rs2.EOF then
KeyID = rs2("KeyID")
InfSubj = rs2("InfSubj")
LogAUser = Show_Text(rs2("LogAUser")&"")
  If LogAUser="" Then 
    LogAUser="(匿名)"
  End If
LogATime = rs2("LogATime")
InfSubj = "<a href='bview.asp?ID="&KeyID&"' target='_self'>"&InfSubj&"</a>"
else
LogAUser = "(无)"
LogATime = "(无)"
InfSubj = "(无)"
end if
  rs2.close()
  set rs2 = nothing

 sNam2 = ""
 if TypNam2<>"" Then
   aNam2 = Split(TypNam2,",")
   for i=0 To uBound(aNam2)
	 sNam2 =sNam2&"<a href='blist.asp?TypID=($User$)&sUser="&aNam2(i)&"' target='_blank'>"&aNam2(i)&"</a>"
     if i<>uBound(aNam2) then 
     sNam2 =sNam2&"<br>"
     end if
   next
 else
     sNam2 ="(无)" 
 end if	
%>
          <tr>
            <td align="center" bgcolor="#FFFFFF"><IMG height=22 src="bimg/<%=nImg%>.gif" width=20> <br /></td>
            <td valign="top" bgcolor="#FFFFFF"><A href="blist.asp?TypID=<%=TypID%>&TypLay=<%=TypLayer%>"><FONT color=#000066>※<%=TypName%>※</FONT></A><br />
              <div class="note02"><%=TypResume%> </div></td>
            <!--<td align="center" bgcolor="#FFFFFF"><img height="60" src="<%=ImgName%>" border="0" /></td>-->
            <td align="center" bgcolor="#FFFFFF"><%=mPub%> / <%=mReply%></td>
            <td bgcolor="#FFFFFF" style="line-height:120%;"><a href="blist.asp?ID=<%=KeyID%>"><%=InfSubj%></a><br />
              by <a href='blist.asp?TypID=($User$)&sUser=<%=LogAUser%>' target='_blank'><%=LogAUser%></a><br />
              <%=LogATime%></td>
            <td align="center" bgcolor="#FFFFFF"><a href="#"> <%=sNam2%></a></td>
          </tr>
          <%
	    rs.movenext
	  loop
	  rs.close()
%>
        </table></td>
    </tr>
    <%End If%>
  </table>
  <div class="line08">&nbsp;</div>
  <%
Next
%>
  <!-- Main BBB /////////////////////////////// -->
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
      <td width="50%" height="24" align="left" bgcolor="#FFFFFF">&nbsp;图例:&nbsp;<img src="bimg/Board2.gif" width="20" height="22" align="absmiddle" /> 有新贴&nbsp;&nbsp;<img src="bimg/Board1.gif" width="20" height="22" align="absmiddle" /> 没新帖</td>
      <td width="50%" align="left" bgcolor="#FFFFFF">提示:
        <%If bbsUser&""<>"" Then%>
        <%=bbsUser%> 您好，您已经登陆系统! &nbsp; <a href="<%=bbsUPass%>" target="_blank">会员中心</a> 或 <a href="<%=bbsLogin%>?send=out&goPage=<%=bbsPath%>">安全登出</a>
        <%Else%>
        您还没有登陆系统 &nbsp; <a href="<%=bbsLogin%>?goPage=<%=bbsPath%>">登录系统</a>
        <%End If%>
      </td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
  <table width="960" border="0" align="center" cellpadding="1" cellspacing="1" class="sysTabC1">
    <tr>
      <td height="60" align="left" valign="top" class="sysTabGB"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="sysTabC2">
          <tr>
            <th width="33%" align="left" bgcolor="#FFFFFF" class="sysTopBG fntFFF" >统计信息</th>
            <th width="33%" align="left" bgcolor="#FFFFFF" class="sysTopBG fntFFF">个人信息</th>
            <th width="33%" align="left" bgcolor="#FFFFFF" class="sysTopBG fntFFF">友情链接</th>
          </tr>
          <tr>
            <td align="left" bgcolor="#FFFFFF"><%
nAllOrg = rs_Count(conn," BBSInfo WHERE KeyMod='"&ModID&"' AND LEN(KeyRE)<12 ")
nAllRep = rs_Count(conn," BBSInfo WHERE KeyMod='"&ModID&"' AND LEN(KeyRE)>12 ")
nDay = Date()
nDayOrg = rs_Count(conn," BBSInfo WHERE KeyMod='"&ModID&"' AND LogATime>#"&nDay&"# AND LEN(KeyRE)<12 ")
nDayRep = rs_Count(conn," BBSInfo WHERE KeyMod='"&ModID&"' AND LogATime>#"&nDay&"# AND LEN(KeyRE)>12 ")

 sql = "SELECT TOP 1 * FROM [BBSInfo] "
 sql=sql& " WHERE KeyMod='"&ModID&"' AND SetShow='Y' AND LEN(KeyRE)<12 ORDER BY KeyID DESC " 
 rs.Open sql,conn,1,1 
 If Not rs.EOF Then
  NewKeyID = rs("KeyID") 
  NewSubj = rs("InfSubj") 
  NewUser = rs("LogAUser") 
  If NewUser="" Then 
    NewUser="(匿名)"
  End If
 Else
  NewKeyID = KeyID
  NewSubj = InfSubj 
  NewUser = "(Nul)" 
 End If
 rs.Close()
%>
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td>总帖子数：<%=nAllOrg%>篇</td>
                  <td>总回复数：<%=nAllRep%>篇 </td>
                </tr>
                <tr>
                  <td>今日帖子：<%=nDayRep%>篇</td>
                  <td>今日回复：<%=nDayRep%>篇 </td>
                </tr>
                <tr>
                  <td colspan="2">最新发表： <a href="blist.asp?ID=<%=NewKeyID%>"><%=NewSubj%></a></td>
                </tr>
                <tr>
                  <td colspan="2">发表会员： <a href='blist.asp?TypID=($User$)&sUser=<%=NewUser%>' target='_blank'><%=NewUser%></a></td>
                </tr>
              </table></td>
            <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><%If bbsUser&""<>"" Then%>
                    我已经登陆： <a href="buser.asp" target="_blank">查看我的帖子</a>
                    <%Else%>
                    您还没有登陆系统 &nbsp; <a href="<%=bbsLogin%>?goPage=<%=bbsPath%>">登录系统</a>
                    <%End If%>
                  </td>
                </tr>
                <tr>
                  <td>我的IP地址：<%=Get_CIP()%> </td>
                </tr>
                <tr>
                  <td>我的操作系统： <%=GetCurrOS()%></td>
                </tr>
                <tr>
                  <td>我的浏览器： <%=GetCurrBrs()%></td>
                </tr>
              </table></td>
            <td align="left" valign="top" bgcolor="#FFFFFF"><!--#include file="../upfile/sys/para/bbslink.htm"-->
            </td>
          </tr>
        </table></td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<!--#include file="_ibot.asp"-->
</body>
</html>
