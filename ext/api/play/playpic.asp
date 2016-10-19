
<%If rpPlayer="bcastr" Then%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="<%=rp1Width%>" height="<%=rp1SwfH%>">
  <param name="movie" value="../ext/api/play/playpid.swf">
  <param name="wmode" value="transparent">
  <param name="FlashVars" value="bcastr_file=<%=rp1Img%>&bcastr_link=<%=rp1Url%>&bcastr_title=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>">
  <embed src="../ext/api/play/playpid.swf" FlashVars="bcastr_file=<%=rp1Img%>&bcastr_link=<%=rp1Url%>&bcastr_title=<%=rp1Txt%>" bgcolor="#FFFFFF" quality="high" width="<%=rp1Width%>" height="<%=rp1SwfH%>" type="application/x-shockwave-flash" />
</object>
<%ElseIf rpPlayer="playp02" Then%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="<%=rp1Width%>" height="<%=rp1SwfH%>">
  <param name="movie" value="../ext/api/play/playp02.swf">
  <param name="wmode" value="transparent">
  <param name="FlashVars" value="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>">
  <embed src="../ext/api/play/playp02.swf" FlashVars="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>" bgcolor="#FFFFFF" quality="high" width="<%=rp1Width%>" height="<%=rp1SwfH%>" type="application/x-shockwave-flash" />
</object>
<%ElseIf rpPlayer="playp03" Then%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="<%=rp1Width%>" height="<%=rp1SwfH%>">
  <param name="movie" value="../ext/api/play/playp03.swf">
  <param name="wmode" value="transparent">
  <param name="FlashVars" value="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>">
  <embed src="../ext/api/play/playp03.swf" FlashVars="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>" bgcolor="#FFFFFF" quality="high" width="<%=rp1Width%>" height="<%=rp1SwfH%>" type="application/x-shockwave-flash" />
</object>
<%ElseIf Request("_rpPlayer")="_Demo.code.asp" Then%>
<%
' AND SetHot='Y'
'sql = "SELECT TOP 60 * FROM [InfoNews] WHERE KeyMod LIKE 'InfN2%' AND LEN(ImgName)>15 ORDER BY LogATime DESC " 
sql = "SELECT TOP 6 * FROM [InfoNews] WHERE KeyMod LIKE 'InfN2%' AND LEN(ImgName)>15 AND SetHot='Y' ORDER BY LogATime DESC " 
rp1Arr = get_Player(sql,"iview.asp?KeyID=($KeyID)")			

rp1Img=rp1Arr(0)
rp1Url=rp1Arr(1)
rp1Txt=rp1Arr(2)
'-------------------------
rp1Width = 220
rp1Height = 170
rp1TxtH = 0
rp1SwfH = rp1TxtH+rp1Height
rpPlayer = "bcastr" ':bcastr,playp02,null
%>
<%ElseIf Request("_rpPlayer")="_Show.bcastr.xml" Then%>
<?xml version="1.0" encoding="utf-8"?>
<bcaster autoPlayTime="3">
 <item item_url="../../../upfile/myfile/logo/logo-2008.jpg" link="http://www.com.cn" itemtitle="test001-Title"></item>  
 <item item_url="../../../upfile/myfile/logo/logo-2009.jpg" link="http://www.com.cn" itemtitle="test001-Title2"></item>  
 <item item_url="../../../upfile/myfile/logo/logo-2010.jpg" link="http://www.com.cn" itemtitle="test001-Title3"></item>  
</bcaster>
<%ElseIf Request("_rpPlayer")="_Test.playpid.swf" Then%>
<!-- Test ////////////////////////////////// -->
/playpid.swf 使用xml文件:<br>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="120" height="120">
  <param name="movie" value="playpid.swf?bcastr_xml_url=playpic.asp?_rpPlayer=_Show.bcastr.xml">
  <param name="wmode" value="transparent">
  <embed src="playpid.swf?bcastr_xml_url=playpic.asp?_rpPlayer=_Show.bcastr.xml" width="120" height="120" type="application/x-shockwave-flash"></embed>
</object>
<br>xml文件<br>
<?xml version="1.0" encoding="utf-8"?>
<bcaster autoPlayTime="3">
 <item item_url="../../../upfile/myfile/logo/logo-2008.jpg" link="http://www.com.cn" itemtitle="test001-Title"></item>  
 <item item_url="../../../upfile/myfile/logo/logo-2009.jpg" link="http://www.com.cn" itemtitle="test001-Title2"></item>  
 <item item_url="../../../upfile/myfile/logo/logo-2010.jpg" link="http://www.com.cn" itemtitle="test001-Title3"></item>  
</bcaster>
<!-- Test ////////////////////////////////// -->
<%Else%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="<%=rp1Width%>" height="<%=rp1SwfH%>">
  <param name="movie" value="../ext/api/play/playpic.swf">
  <param name="wmode" value="transparent">
  <param name="FlashVars" value="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>">
  <embed src="../ext/api/play/playpic.swf" FlashVars="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>" bgcolor="#FFFFFF" width="<%=rp1Width%>" height="<%=rp1SwfH%>" type="application/x-shockwave-flash" />
</object>
<%End If%>