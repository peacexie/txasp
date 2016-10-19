<!--#include file="rss_config.asp"-->
<%

rPath = Request.ServerVariables("URL")
rPath = Replace(rPath,"rss.asp","") '' Get_vPath(7)
''RssID,Title,TabID,ModID,TypID,ListUrl,ViewUrl
aUrlLI(i) = Replace(aUrlLI(i),"(ModID)",aModID(i))
aUrlLI(i) = Replace(aUrlLI(i),"(TypID)",aTypID(i))
aUrlVW(i) = Replace(aUrlVW(i),"(ModID)",aModID(i))
aUrlVW(i) = Replace(aUrlVW(i),"(TypID)",aTypID(i))

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE>RSS订阅列表-<%=Config_Name%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../../pfile/pimg/style.css">
</HEAD>
<BODY>

<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td width="10"><img height="42" src="../../pfile/pimg/qqimg1_06_01.jpg" width="10" /></td>
    <td width="120" align="center" background="../../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle">RSS订阅列表</td>
    <td width="10"><img height="42" src="../../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="210" height="30" align="left" class="SysTit02">&nbsp;</td>
          <td align="right" class="SysTit03"><%=vPMsg_WSite%><%=vPMsg_WHome%> &gt;&gt; RSS订阅列表 &nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;">
    <div style="line-height:12px;">&nbsp;</div>
      <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td height="30" colspan="2" align="left" class="SysCont" style="padding:5px;"><span class="SysCont" style="padding:5px;"><img src="../../img/tool/rss.jpg" width="45" height="47" hspace="8" align="left" /></span>　　欢迎订阅 <%=Config_Name%> RSS，请将您需要的RSS Feeds的XML地址添加到您的RSS阅读器即可(点击RSS图标生成RSS   Feeds的XML地址)。 </td>
        </tr>
        <tr>
          <td width="210" height="30" align="center" background="../../pfile/pimg/qqimg21_12_01.gif" class="SysTitle">RSS订阅列表</td>
          <td width="500" align="right" nowrap="nowrap" background="../../pfile/pimg/qqimg21_12_02.gif">&nbsp;</td>
        </tr>
        <tr>
          <td height="120" colspan="2" align="left" valign="top" style="padding:8px 8px;"><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr>
              <th align="center" bgcolor="#FFFFFF">网站源信息</th>
              <th align="center" bgcolor="#FFFFFF">RSS地址</th>
              <th align="center" bgcolor="#FFFFFF">RSS订阅</th>
            </tr>
<%
For i=0 To uBound(aRssID)
''RssID,Title,TabID,ModID,TypID,ListUrl,ViewUrl
aUrlLI(i) = Replace(aUrlLI(i),"(ModID)",aModID(i))
aUrlLI(i) = Replace(aUrlLI(i),"(TypID)",aTypID(i))
aUrlVW(i) = Replace(aUrlVW(i),"(ModID)",aModID(i))
aUrlVW(i) = Replace(aUrlVW(i),"(TypID)",aTypID(i))
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F4F4F4"
		End If
%>

            <tr bgcolor="<%=col%>">
              <td><a href="<%=aUrlLI(i)%>"><%=aTitle(i)%></a></td>
              <td><%=rPath%>rss_list.asp?TypID=<%=aRssID(i)%></td>
              <td align="center"><a href="rss_list.asp?TypID=<%=aRssID(i)%>">RSS订阅</a></td>
            </tr>
<%
Next
%>
            </table>
          </td>
        </tr>
    </table></td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240"><!-- #include file="rss_side.asp" --></td>
  </tr>
</table>

</BODY>
</HTML>
