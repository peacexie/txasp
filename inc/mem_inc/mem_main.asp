<!--#include file="../../sadm/func2/func_const.asp"-->
<!--#include file="../../sadm/func2/func_perm.asp"-->
<%Response.CodePage=65001%>
<%Response.Charset="UTF-8"%>
<%
varDef = "1"
verMemb = Request("verMemb")
If verMemb="" Then
  verMemb = varDef
End If
%>
<!--#include file="../../pfile/lang/vmemb.asp"-->
<%

'Response.Write verMemb
rPage = Request("rPage")
rMenu = Request("rMenu")
If Session("rPage")&""<>"" Then
 rPage=Session("rPage")
ElseIf rPage<>"" Then
 rPage=rPage
Else
 rPage="../../member/info/index.asp?verMemb="&verMemb '" '?verMemb="&verMemb
End If
If rMenu<>"" Then
   rMenu="UsrMenu('"&rMenu&"');"
End If

%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="robots" content="noindex, nofollow">
<title><%=Session("MemID")%>-<%=vMADM_Title%>-<%="["&vPMsg_WName&"]"%></title>
<link rel="stylesheet" href="../../inc/mem_img/default.css" type="text/css" />
</head>
<body>
<table width="100%" height="100%" id="IndexTableBody" cellpadding="0">
  <thead>
    <tr>
      <th width="180" style="width:180px; overflow:hidden"> <div id="logoMemb1"></div></th>
      <th height="10"> <div id="logoMemb2"></div> <A 
href="?verMemb=1" id="VerSetC1" onClick="SetLang('GB2312')" target="_self">转简体</A> | <a 
href="?verMemb=2" target="_self">English</A> | <a 
href="?verMemb=0" id="VerSetC0" onClick="SetLang('Big5')" target="_self">轉繁體</A>
    </th>
    </tr>
  </thead>
  <tr>
    <td height="10" class="menu"><ul class="bigbtu">
        <div id="Pea01"><%=Left(Session("MemID")&"",15)%></div>
      </ul></td>
    <td height="10" class="tab"><ul id="TabPage1">
        <li id="Tab1" class="Selected"><span id="SubTilte"><%=vMADM_PAct%></span></li>
        <li id="Tab2" style="display:block" onClick="RefMainFrame(-1)"><span style="display:block"><%=vMADM_PBack%></span></li>
        <li id="Tab2" style="display:block" onClick="RefMainFrame(0)" ><span style="display:block"><%=vMADM_PLoad%></span></li>
        <li id="Tab2" style="display:block" onClick="RefMainFrame(1)" ><span style="display:block"><%=vMADM_PForw%></span></li>
        <!--<li id="Tab3"><span id="dnow99" style="display:block">空白页</span></li>-->
      </ul></td>
  </tr>
  <tr>
    <td class="t1"><div id="contents">
        <table cellpadding="0">
          <tr class="t1">
            <td><div class="menu_top"></div></td>
          </tr>
          <tr class="t2">
            <td><!--#include file="mem_menu.asp"--></td>
          </tr>
          <tr class="t3">
            <td><div class="menu_end"></div></td>
          </tr>
        </table>
      </div></td>
    <td class="t2">
        <div style="height:100%; border:1px solid #03C">
          <iframe src="<%=rPage%>" name="UsrMain" frameborder="0" scrolling="yes" width="100%" height="100%"></iframe>
        </div>
    </td>
  </tr>
</table>
<%Session("rPage")=""%>
<script type="text/javascript">
function RefMainFrame(xPage)
{
  self.parent.frames["UsrMain"].history.go(xPage);
}
</script>
<script type="text/javascript">if(top.location!==self.location){top.location=self.location;}</script>
<script src='../../ext/api/conv/convert.js' type="text/javascript"></script>
<%
scr = "script"
If verMemb="0" Then
 Response.Write "<"&scr&" type='text/javascript'>SetLang('Big5');</"&scr&">"
Else 'If verMemb="1" Then
 Response.Write "<"&scr&" type='text/javascript'>SetLang('xOrg');</"&scr&">"
End If
%>
</body>
</html>
