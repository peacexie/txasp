<!--#include file="binc/_config.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<%
TypID = RequestS("TypID","C",48)
TypLay = RequestS("TypLay","C",255)
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)

If KW&"" <> "" Then
 sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
 sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
 sqlK = sqlK & " ) " 
End If
If TypID="($User$)" Then
 sUser = RequestS("sUser","C",24)
 sqlK = " AND ( LogAUser='"&sUser&"') "
 TName = "["&sUser&"] 的帖子"
ElseIf TypID="($Self$)" Then '($Else$)
 sUser = RequestSafe(bbsUser,"C",24)
 sqlK = " AND ( LogAUser='"&sUser&"') "
 TName = "["&sUser&"] 的帖子"
ElseIf TypID="($Search$)" Then
 If TypLay&"" <> "" Then
   sqlK = sqlK & " AND ( InfType='"&TypLay&"') " 
 End If
 sUser = RequestS("sUser","C",24)
 If sUser&"" <> "" Then
   sqlK = " AND ( LogAUser='"&sUser&"') "
 End If
 TName = "…帖子搜索…"
ElseIf TypID&"" <> "" Then
 sqlK = sqlK & " AND ( InfType='"&TypLay&"') " 
 TName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&TypID&"'")
End If
If Request("SetHot")="Y" Then
 sqlK = sqlK & " AND ( SetHot='Y') "
End If
    
%>
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
      <td align="left" bgcolor="#FFFFFF">&nbsp;<img src="bimg/face1.gif" align="absmiddle" />&nbsp;<a href="../">首页</a> &gt;&gt; <a href="bind.asp"><%=bbsName%></a> &gt;&gt; <%=TName%> &gt;&gt; </td>
      <td width="50%" align="right" bgcolor="#FFFFFF"><a href="badd.asp?TypID=<%=TypID%>&TypLay=<%=TypLay%>"><img src="bimg/postnew.gif" width="106" height="23" border="0" align="absmiddle" /></a> &nbsp;<a href="blist.asp?TypID=<%=TypID%>&TypLay=<%=TypLay%>&amp;SetHot=Y"><img src="bimg/jinghua1.gif" width="78" height="23" border="0" align="absmiddle" /></a> &nbsp;&nbsp;&nbsp;</td>
    </tr>
  </table>
  <%

    sql = " SELECT * FROM [BBSInfo] "
	sql =sql& " WHERE KeyMod='"&ModID&"' AND SetShow='Y' AND LEN(KeyRE)<12 "
	sql =sql& " "&sqlK
	sql =sql& " ORDER BY KeyID DESC" 
   'Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 8 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5B8DC9">
    <tr align="left" bgcolor="#FFFFFF" style="word-break:break-all; border-top:1px solid #CCCCCC; line-height:150%;">
      <td align="center" class="sysTopBG fntFFF">状态</td>
      <td width="50%" align="center" class="sysTopBG fntFFF">主题</td>
      <td align="center" nowrap="nowrap" class="sysTopBG fntFFF">作者</td>
      <td align="center" nowrap="nowrap" class="sysTopBG fntFFF">点击/回复</td>
      <td align="center" nowrap="nowrap" class="sysTopBG fntFFF">时间</td>
    </tr>
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize

KeyID = rs("KeyID")
InfSubj = Left(rs("InfSubj"),120)
InfType = rs("InfType")
InfCont = Show_Text(rs("InfCont"))
LnkName = Show_Text(rs("LnkName"))
SetRead = rs("SetRead")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
nReply = rs_Count(conn," BBSInfo WHERE KeyRE='"&KeyID&"' ")
nDays = DateDiff("d",LogATime,Date)
	  %>
    <tr align="left" bgcolor="#FFFFFF">
      <td align="center"><%if SetHot="Y" then%>
        <img src="../img/tool/icon_jian.gif" width="15" height="15" align="absmiddle">
        <%elseif cStr(nDays)="0" then%>
        <img src="../img/tool/new2.gif" width="30" height="10" align="absmiddle">
        <%else%>
        <img src="../img/tool/email2.gif" width="16" height="16" align="absbottom">
        <%end if%></td>
      <td width="50%" align="left"><A href="bview.asp?ID=<%=KeyID%>"><%=InfSubj%></A></td>
      <td align="center"><%=LogAUser%></td>
      <td align="center"><%=SetRead%>/<%=nReply%></td>
      <td align="center"><%=LogATime%></td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
    <tr align="center">
      <td height="27" colspan="5" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send=pag&TypID="&TypID&"&TypLay="&TypLay&"&KW="&KW&"",1)%></td>
    </tr>
    <%  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="5">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  'Set rs = Nothing
	  
	  %>
  </table>
  <div style="line-height:8px;">&nbsp;</div>
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
      <td height="24" align="center" bgcolor="#FFFFFF">&nbsp;<img src="../img/tool/email2.gif" width="16" height="16" align="absmiddle" /> 普通帖子&nbsp;&nbsp;&nbsp;<img src="../img/tool/icon_jian.gif" width="15" height="15" align="absmiddle" /> 推荐帖子&nbsp;&nbsp;&nbsp;<img src="../img/tool/new2.gif" width="30" height="10" align="absmiddle" /> 今日新帖&nbsp;</td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<!--#include file="_ibot.asp"-->
</body>
</html>
