<!--#include file="../page/_config.asp"-->
<%
Page = RequestS("Page","N",1)
MD = RequestS("MD","C",24) : If MD="" Then MD="BBSVD24"
TM = RequestS("TimID","C",12) 'All,Now,Wait,End
sqlK=" KeyMod='"&MD&"' "
If TM="All" Then
  
ElseIf TM="Wait" Then
 'sqlK=sqlK&" AND "&Cfg_FTime&Date()&Cfg_FTime&"<InfTime1 "
ElseIf TM="End" Then
 'sqlK=sqlK&" AND InfTime2<"&Cfg_FTime&Date()&Cfg_FTime&" "
Else 'If 'Now
 'sqlK=sqlK&" AND "&Cfg_FTime&Date()&Cfg_FTime&" BETWEEN InfTime1 AND InfTime2 "
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网上调查-<%=vPMsg_WName%></title>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<%
    sql = " SELECT * FROM [VoteInfo] "
	sql =sql& " WHERE " &sqlK& " ORDER BY InfTime1 DESC" 
   rs.Open Sql,conn,1,1 ': Response.Write sql
   rs.PageSize = 20 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<div style="padding:8px;">
<table width="100%"  border="0" align="center" cellpadding="2" cellspacing="1">
  <tr class="fnt666">
    <th height="27" nowrap>&nbsp;</th>
    <th align="left" nowrap>投票主题</th>
    <th width="80" align="center" nowrap>开始时间</th>
    <th width="80" align="center" nowrap>结束时间</th>
    <th width="60" align="center" nowrap>已投票数</th>
  </tr>
  <tr>
    <td colspan="5" nowrap bgcolor="#999999"></td>
  </tr>
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
'TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&KeyMod&"'") 'rs("TypName")
InfSubj = rs("InfSubj")
ImgName = rs("ImgName")
InfSubj = Show_Text(InfSubj) 'Show_SetSubj(InfSubj,SetHot,SetSub)
InfTim1 = rs("InfTime1")
InfTim2 = rs("InfTime2")
InfCard = rs("InfCard")
InfVNum = rs("InfVNum")
InfNum1 = rs("InfNum1")
InfNum2 = rs("InfNum2")
goUrl = "svote.asp"
If DateDiff("d",InfTim2,Date())>0 Then
goUrl = "sview.asp"
End If
smVote=rs_Count(conn,"VoteLogs WHERE KeyMod='"&KeyID&"' ")
	  %>
  <tr>
    <td width="3%" height="24" align="center">&middot;</td>
    <td><a href="#" onClick="javascript:showModalDialog('<%=goUrl%>?ID=<%=KeyID%>','x','center=yes;dialogWidth=640px;dialogHeight=480px');"><%=InfSubj%></a></td> 
    <td align="center" nowrap>&nbsp;&nbsp;&nbsp;<%=InfTim1%></td>
    <td align="center" nowrap><%=InfTim2%>&nbsp;&nbsp; &nbsp;</td>
    <td align="center" nowrap>&nbsp;&nbsp;&nbsp;<%=smVote%> 人次</td>
  </tr>
  <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
  <tr>
    <td colspan="5" nowrap bgcolor="#999999"></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="5" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send="&send&"&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"&SubID="&SB&"",1)%></td>
  </tr>
  <%  
  Else
  %>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="5">无信息</td>
  </tr>
  <%
  End If
	  
	  rs.Close()
	  
	  %>
</table>
</div>
</BODY>
</HTML>
