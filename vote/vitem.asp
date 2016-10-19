<!--#include file="_config.asp"-->
<%

ID = RequestS("ID","C",48)

  rs.Open "SELECT * FROM [VoteItem] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfRem = Show_Html(rs("InfRem")) 
SetVote = rs("SetVote")
ImgName = rs("ImgName")
If ImgName<>"" Then
 sImg = "../upfile/vote/"&ImgName&""
Else
 sImg = "../img/logo/no_pic211.jpg"
End If
  end if
  rs.Close()

  rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&KeyMod&"'",conn,1,1 
  if NOT rs.eof then 
bInfSubj = Show_Text(rs("InfSubj"))
bInfTim1 = rs("InfTime1") 
bInfTim2 = rs("InfTime2") 
  end if 
  rs.Close()

'Response.Write ID&KeyMod
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线调查投票系统-<%=Config_Name%></title>
<style type="text/css">
<!--
.vH100 { line-height:100%; }
body {
	margin-left: 0px;
	margin-top: 8px;
	margin-right: 0px;
	margin-bottom: 8px;
}
body,td,th {
	font-size: 13px;
}
-->
</style>
<script src="vote.js" type="text/javascript"></script>
</HEAD>
<BODY>
<table width="570" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
      <tr>
        <td height="30" colspan="2" align="center" bgcolor="#FFFFFF"><strong><%=Config_Name%> 在线调查投票系统</strong></td>
        </tr>
      <tr>
        <td width="30%" align="center" bgcolor="#FFFFFF">选项名称</td>
        <td bgcolor="#FFFFFF"><%=InfSubj%></td>
      </tr>
      <tr>
        <td width="30%" align="center" bgcolor="#FFFFFF">所的票数</td>
        <td bgcolor="#FFFFFF"><%=SetVote%> 票&nbsp; </td>
      </tr>
      <tr>
        <td width="30%" align="center" bgcolor="#FFFFFF">投票项目</td>
        <td bgcolor="#FFFFFF"><%=bInfSubj%></td>
      </tr>
      <tr>
        <td width="30%" align="center" bgcolor="#FFFFFF">起止时间</td>
        <td bgcolor="#FFFFFF"><%=bInfTim1%> ~ <%=bInfTim2%></td>
      </tr>

    </table></td>
    <td width="150" height="120" align="center"><a 
 href="<%=sImg%>"><img 
 src="<%=sImg%>" alt="Image" width="110" height="90" vspace="1" border="0" 
 onload="javascript:setImgSize(this);"></a></td>
  </tr>
  <tr>
    <td colspan="2" align="left" style="padding:3px;"><div style="background-color:E0E0E0; padding:0 8px; line-height:150%; border:5px #FFFFFF;"><%=InfRem%></div></td>
  </tr>
</table>
<% Set rs = Nothing %>
</BODY>
</HTML>
