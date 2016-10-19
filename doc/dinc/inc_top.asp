<link rel="stylesheet" type="text/css" href="style.css">
<div class="line01">&nbsp;</div>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="sysTopBG">
  <tr>
    <td width="180" align="center"><img src="../pfile/pub/admin.jpg" width="120" height="60" alt="" /></td>
    <td height="80" align="left" valign="top"><div style="font-size:18px; color:#FFF; padding:18px 2px 24px 5px; font-weight:bold"><%=sysName%></div>
      <div style="font-size:12px; color:#00F"> &nbsp;</div></td>
    <td width="300" align="right" valign="top"><div style="font-size:14px; color:#FFF; padding:5px 8px 5px 5px;"> &nbsp; <a href="../bbs/" class="cWhite">内部论坛</a> &nbsp; <a href="../ext/login.asp?send=out" class="cWhite">退出公文</a> </div>
      <div style="font-size:12px; color:#00F; padding:30px 8px 2px 5px;"> 当前用户：<%=Session("InnID")%>
        <IFRAME align="middle" name=RefFrame src="../tools/out/check.asp" frameBorder=0 width="20" scrolling=no height="16"></IFRAME>
        &nbsp; <a href="userpw.asp">修改密码</a> </div></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" id="docMenu" class="sysTopBG mTopRight">
  <tr>
    <td id="mTopRPub" class="mTopComm"><a href="doc_get.asp?Flag=Public">公开公文</a></td>
    <td id="mTopRFlg" class="mTopComm"><a href="doc_get.asp?Flag=NoRead">未读公文</a></td>
    <td id="mTopList" class="mTopComm"><a href="doc_get.asp">定向接收</a></td>
    <td width="1" align="center"></td>
    <td id="mTopPub" class="mTopComm"><a href="info_add.asp">发布公文</a></td>
    <td id="mTopAdm" class="mTopComm"><a href="info_list.asp">管理公文</a></td>
    <td width="1" align="center"></td>
    <td id="mTopHome" class="mTopComm"><a href="index.asp">公文首页</a></td>
    <td id="mTopUser" class="mTopComm"><a href="userpw.asp">会员中心</a></td>
    <td height="25" align="center" class="mTopLeft">&nbsp;</td>
    <td id="mTopOut" class="mTopComm"><a href="../ext/login.asp?send=out" class="cRed">退出公文</a></td>
  </tr>
</table>
