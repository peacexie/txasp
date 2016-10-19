<div class="line01">&nbsp;</div>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="sysTopBG">
  <tr>
    <td width="300" align="center" class="sysTopBG"><a href="bind.asp">
    <img src="bimg/sys-logo.jpg" width="280" height="80" border="0" alt="和平鸽论坛" />
    </a></td>
    <td height="80" align="right" class="sysTopBG" style="padding-right:12px;">
    <!--#include file="../upfile/sys/para/BBSBan1.htm"--></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="sysTopBG">
  <tr>
    <td width="20%" align="center" class="fntFFF"><%=bbsName%></td>
    <td height="21" align="center" class="sysTopBG">&nbsp;</td>
    <td width="60%" align="right" class="fntFFF"> |
    <A href="bsearch.asp" class="LnkWhite">搜索</A> | 
    <%If bbsUser&""<>"" Then%>
    <a href="buser.asp" class="LnkWhite">我的帖子</a> | 
    <a href="<%=bbsUPass%>" target="_blank" class="LnkWhite">会员中心</a> |
    <a href="<%=bbsLogin%>?send=out&goPage=<%=bbsPath%>" class="LnkWhite">登出</a> | 
    <%Else%>
    <a href="<%=bbsLogin%>?goPage=<%=bbsPath%>" class="LnkWhite">登录</a> | 
    <%End If%>
    | <A href="../member/mu_<%=AppRand12%>.asp?goPage=<%=bbsPath%>" target="_blank" class="LnkWhite">注册</A> | 
    <a href="../member/get_pw.asp?goPage=<%=bbsPath%>" target="_blank" class="LnkWhite">忘记密码</a> |
    <a href="../tools/help/uhelp.asp" target="_blank" class="LnkWhite">帮助</a> | &nbsp; </td>
  </tr>
</table>