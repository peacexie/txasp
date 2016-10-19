<!--#include file="../himg/tconfig.asp"-->
<!--#include file="oConfig.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>在线统计</title>
<link rel="stylesheet" type="text/css" href="../himg/tstyle.css">
</head>
<body>
<script src="Online.asp" type="text/javascript"></script>
<table width="620" height="55" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#094C8D">
  <tr bgcolor="#FFFFFF">
    <td width="25%">&nbsp;总访问量: <span id="PeaceAllView">(0)</span></td>
    <td width="25%">&nbsp;在线人数: <span id="PeaceOnlNow">(0)</span></td>
    <td width="25%">&nbsp;昨天访问: <span id="PeaceDatePrev">(0)</span></td>
    <td colspan="2">&nbsp;今天访问: <span id="PeaceDateThis">(0)</span></td>
  </tr>

  <tr bgcolor="#FFFFFF">
    <td colspan="5">&nbsp;最高在线: <span id="PeaceOnlMax">(0)</span> (发生在: <span id="PeaceOnlMTime">(0)</span>) &nbsp;&nbsp;其它外部统计排名参考: <a 
					  href="http://alexa.chinaz.com/" target="_blank">[1]中国站长站</a>, <a 
					  href="http://www.alexa.com/" target="_blank">[2]Alexa官方排名</a></td>
  </tr>
</table>
<br style="line-height:15px; ">
<table width="620" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#094C8D">
  <tr align="center" bgcolor="#FFFFFF">
    <th nowrap>ID</th>
    <th nowrap>IP</th>
    <th nowrap>最新更新</th>
    <th width="50%" nowrap>最近浏览页面</th>
  </tr>
  <%
					
    sql = " SELECT * FROM [OnlPage] "
	sql = sql& " WHERE pTime>="&cfgTimeC&""&DateAdd("n",(-2)*pExpTime,Now())&""&cfgTimeC&" "
	sql = sql& " ORDER BY pTime DESC " 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1 ':Response.Write sql

%>
  <%
	  i = 0
	  Do While Not rs.eof 
	    i = i + 1
		pID = rs("pID")
		pIP = rs("pIP")
		pPage = rs("pPage")
		pTime = rs("pTime") 
		aPag = Split(pPage,";")
		aMax = uBound(aPag)
		j=0 : s="" : k=0
		If aMax>12 Then  '12
		  j = aMax - 12
		End If
		For j=j To uBound(aPag)
		If aPag(j)<>"" Then
		  k=k+1
		  s=s& "<a href='"&aPag(j)&"' target='_blank'>页面"&k&"</a> "
		  If k=7 Then s=s& "<BR>"
		End If
		Next
		col="#F8F8F8" : k=""
		if i mod 2 = 1 then
		  col = "#ffffff"
		end if
		if inStr(pID,Session.SessionID)>0 Then
		  col = "#ffffee"
		  k="<br><font color=red>(我自己)</font> "
		end if
	  %>
  <tr align="center" bgcolor="<%=col%>">
    <td nowrap>&nbsp;<%=i%>
    <!--<%=pID%>--></td>
    <td nowrap><%=pIP%><%=k%></td>
    <td nowrap><%=pTime%></td>
    <td align="left"><%=s%></td>
  </tr>
  <%

	    rs.movenext
	  loop
	  rs.close()
	  set rs = nothing

%>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="7" nowrap>&nbsp;<span class="fntF00">注意:</span> 在线人数为 大约[<%=pExpTime%>]分钟内的统计数据,本列表为 大约[<%=pExpTime*2%>]分钟内的历史记录;</td>
  </tr>
</table>
<script charset='utf-8' type='text/javascript' src='oInfo.asp?send=View01'></script>