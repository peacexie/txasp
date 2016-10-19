<!--#include file="../../page/_config.asp"-->
<div id="DateTime"><script type="text/javascript">setInterval("document.getElementById('DateTime').innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt(new Date().getDay());",1000);</script></div
<%	
'AND InfType LIKE 'S110012%' AND SetHot='Y'	  
sql = "SELECT TOP 4 * FROM [InfoPics] WHERE KeyMod='PicS124' AND LEN(ImgName)>15 ORDER BY SetTop,LogATime DESC " 
rs.Open sql,conn,1,1 
rp1Img="":rp1Url="":rp1Txt=""
Do While NOT rs.EOF
 sID = rs("KeyID")
 sSubj = Show_SLen(rs("InfSubj"),120)
 sSubj = Replace(sSubj,"'","")
 sImg  = get_1Img(sID,rs("ImgName")) 
 rp1Img = rp1Img&sImg&"|"
 rp1Url = rp1Url&"iview.asp?KeyID="&sID&"|"
 rp1Txt = rp1Txt&sSubj&"|"
rs.MoveNext
Loop
rs.Close()

If rp1Url<>"" Then
  rp1Img=rp1Img&"$" : rp1Img=Replace(rp1Img,"|$","")
  rp1Url=rp1Url&"$" : rp1Url=Replace(rp1Url,"|$","")
  rp1Txt=rp1Txt&"$" : rp1Txt=Replace(rp1Txt,"|$","")
End If
%>

<script type="text/javascript">
var nplayer_focus_width=210;
var nplayer_focus_height=180;
var nplayer_text_height=20;
var nplayer_swf_height=nplayer_focus_height+nplayer_text_height ;
var nplayer_pics_list ='<%=rp1Img%>' ;
var nplayer_links_list='<%=rp1Url%>';
var nplayer_texts_list='<%=rp1Txt%>' ;
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ nplayer_focus_width +'" height="'+ nplayer_swf_height +'">');
  document.write('<param name="movie" value="../../inc/home/playpic.swf"><param name="quality" value="high"><param name="bgcolor" value="#ffffff">');
  document.write('<param name="menu" value="false"><param name=wmode value="TRANSPARENT">'); //opaque
  document.write('<param name="FlashVars" value="pics='+nplayer_pics_list+'&links='+nplayer_links_list+'&texts='+nplayer_texts_list+'&borderwidth='+nplayer_focus_width+'&borderheight='+nplayer_focus_height+'&textheight='+nplayer_text_height+'">');
  document.write('<embed src="../../inc/home/playpic.swf" wmode="opaque" FlashVars="pics='+nplayer_pics_list+'&links='+nplayer_links_list+'&texts='+nplayer_texts_list+'&borderwidth='+nplayer_focus_width+'&borderheight='+nplayer_focus_height+'&textheight='+nplayer_text_height+'" menu="false" bgcolor="#ffffff" quality="high" width="'+ nplayer_focus_width +'" height="'+ nplayer_swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
document.write('</object>');  
</script>

<%
Set rs=Nothing
%>
    >

