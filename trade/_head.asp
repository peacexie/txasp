<link href="inc/style.css" rel="stylesheet" type="text/css">
<div class="line05">&nbsp;</div>

<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" class="bgLine1">
  <tr>
    <td width="180" align="center" nowrap="nowrap" id="DateTime"><script type="text/javascript" src="../inc/home/jsdat1.js"></script></td>
    <td align="center">&nbsp;</td>
    <td width="420" height="30" align="center" nowrap="nowrap">
      <li class="SysVers" id="SysEng"><a href="ypage.asp">会员黄页</a></li>
      <li class="SysVers" id="SysEng"><a href="info.asp?ModID=TraT124">会员供求</a></li>
      <li class="SysVers" id="SysEng"><a href="info.asp?ModID=TraN124">行业新闻</a></li>
      <li class="SysVers" id="SysEng"><a href="info.asp?ModID=TraJ124">企业招聘</a></li>
      <li class="SysVers" id="SysEng"><a href="index.asp"><span class="cRed">供求首页</span></a></li>
    </td>
  </tr>
</table>

<div class="line05">&nbsp;</div>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="150" align="center"><a href="../" target="_parent"><img src="../pfile/pub/logo.jpg" width="90" height="60" border="0" /></a><a href="" target="_parent"></a></td>
    <td nowrap="nowrap">&nbsp;</td>
    <td width="60%" nowrap="nowrap">
      <%If US="" Then%>
      <div id='ImgPlayer_AdSTop' style='width:750px; height:60px; overflow:hidden; border:1px solid #CCC'>
      </div><script src='<%=Config_Path%>upfile/sys/xadv/pic_AdSTop.js'></script>
      <%Else%>
      <div style='width:750px; height:60px; overflow:hidden; border:1px solid #CCC' class="traNam1"><%=MemSubj%> <span class="traNam2">[<%=MemMod%>]</span></div>
      <%End If%>
    </td>
  </tr>
</table>
<div class="line05">&nbsp;</div>
