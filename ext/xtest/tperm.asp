<!--#include file="../_admin/config.asp"-->
<%
If Request("Act")="tstPerm" Then
  rDir = Config_Path&"tools/tools.asp"
  If Chk_URL3(rDir)<>"eUrl" Then
    Session("ChkCode")="tstPerm"&Config_Code
	Response.Redirect "tperm.asp"
  End If 
End If
If Session("ChkCode")<>"tstPerm"&Config_Code Then
  Response.End()
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<base target="_blank" />
<title>权限测试</title>
<style type="text/css">
<!--
body {
	margin:5px;
}
body, td, th {
	font-size: 14px;
}
.bWidth {
	width:720px;
	margin:auto;
	text-align:left;
}
.bBorder {
	border:1px #CCCCCC solid;
}
.bItem {
	width:98%;
	text-align:left;
	border-bottom:1px #CCC solid;
	clear:both;
	line-height:5px;
	margin:3px 0px;
}
.bItm1 {
	width:113px;
	height:20px;
	padding:0px 0px 0px 2px;
	display:inline-block;
	float:left;
	line-height:120%;
	font-weight:bold;
	overflow:visible;
	background-color:#F0F0F0;
}
.bItm3 {
	width:113px;
	height:20px;
	padding:0px 0px 0px 2px;
	display:inline-block;
	float:left;
	line-height:120%;
	overflow:visible;
}
.Fnt00F {
	color:#00F;
	padding:0 3px 0 3px;
}
.FntF00 {
	color:#F00;
	padding:0 3px 0 3px;
}
a:link {
	color: #009;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #009;
}
a:hover {
	text-decoration: underline;
	color: #F0F;
}
a:active {
	text-decoration: none;
	color: #F00;
}
-->
</style>
</head>
<body>
<%
SET objFSO = Server.CreateObject("Scripting.FileSystemObject")

Dim TabDir(8) : nMax = 3
TabDir(0) = "tools/tools.asp;tools/base/omaster.asp"
TabDir(1) = "bbs;doc"
TabDir(2) = "sadm;admin;system;user"
TabDir(3) = "smod;adupd;file;gbook;info;link;type;vote"
TabDir(4) = ""

%>
<div class="bWidth">
  <fieldset style="padding:3px;">
    <legend> Peace </legend>
    <div class='bItm1'>Links</div>
    <div class='bItm3'>?&amp;#</div>
    <div class='bItm3'><%=Server.URLEncode("?&#")%></div>
    <div class='bItm3'><a href="../../upfile/%23dbf%23/ysWeb.asp" target="_blank">ysWeb.asp</a></div>
    <div class='bItm3'><a href="../../tools/base/shome.asp?Act=InjTrap" target="_blank">Act=InjTrap</a></div>
    <div style="clear:both;">&lt;% ahaha call ~!@#$%^&amp;*()_+ b=3-&quot;aabc&quot; if(&lt;&gt;:&quot;{}?&gt;) %&gt; <br />
      &lt;% aa.bb&gt;cc@dd %&gt;<br />
      &lt;% 哈哈，小样，数据库能随便给你下吗？call abasdfc,bcd   b=3-&quot;aabc&quot; %&gt;<br />
      &lt;script type='text/javascript'&gt;<br />
      window.open('../../tools/base/shome.asp?Act=InjTrap');<br />
      &lt;/script&gt;<br />
      &lt;%Response.End()%&gt;</div>
  </fieldset>
  <fieldset style="padding:3px;">
    <legend> Master Center </legend>
    <div style="float:right"> <a href="../../inc/adm_inc/adm_main.asp">Admin</a> | <a href="../../inc/mem_inc/mem_main.asp">Member</a> |
      Logout
      | </div>
    Peace Perm Test Center ... <span class="FntF00">注意，测试后可能要重新登陆！</span>
  </fieldset>
  <!-- 2~N //////////////////////////////////////// -->
  <%
  For i = 2 To nMax
  s1 = TabDir(i)
  a1 = Split(s1,";")
%>
  <fieldset style="padding:3px;">
    <legend> \\<%=a1(0)%>\ </legend>
    <%
	  For j=1 To uBound(a1)
	    s2 = fList(a1(0),a1(j))
	    Response.Write vbcrlf&vbcrlf&"<div class='bItm1'>"&a1(j)&"\</div>"&s2&vbcrlf&"<div class='bItem'>&nbsp;</div>"
      Next
    %>
  </fieldset>
  <%
  Next
%>
  <!-- 0:Files //////////////////////////////////////// -->
  <%
  a2 = Split(TabDir(0),";")
%>
  <fieldset style="padding:3px;">
    <legend> [Files] </legend>
    <%
	  For j=0 To uBound(a2)
	    iFile = Mid(a2(j),InStrRev(a2(j),"/")+1) 
		Response.Write vbcrlf&"<div class='bItm3'><a href='../../"&a2(j)&"'>"&iFile&"</a></div>"
      Next
    %>
  </fieldset>
  <!-- 1:Root //////////////////////////////////////// -->
  <%
  a3 = Split(TabDir(1),";")
%>
  <fieldset style="padding:3px;">
    <legend> [Root] </legend>
    <%
	  For j=0 To uBound(a3)
	    s3 = fList(a3(j),"")
	    Response.Write vbcrlf&vbcrlf&"<div class='bItm1'>"&a3(j)&"\</div>"&s3&vbcrlf&"<div class='bItem'>&nbsp;</div>"
      Next
    %>
  </fieldset>
</div>
<%

Function fList(xP1,xP2)
Dim PN,i,s : s=""
  PN = "../../"&xP1&"/"&xP2&"/"
  PN = Replace(PN,"//","/")
  If fold_exist(PN,"") Then
  SET objFolder = objFSO.GetFolder(Server.MapPath(PN))
  FOR EACH objFile IN objFolder.Files
  iName = objFile.Name 
  iExt = Mid(iName,InStrRev(iName,"."),8) 
    If inStr(".asp.",lCase(iExt))>0 Then
	  li = "<div class='bItm3'><a href='"&PN&objFile.Name&"'>"&objFile.Name&" </a></div>"
	  s = s&vbcrlf&li
    End If
  Next
  SET objFolder = Nothing
  Else
	  s = "<div class='bItm3'> Null </div>"
  End If
fList = s
End Function

SET objFSO = Nothing
%>
</body>
</html>
