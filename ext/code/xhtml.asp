<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Demo Html</title>
</head>
<body>
<hr>
<script type=text/javascript>
var pic_width=950; //图片宽度
var pic_height=60; //图片高度
var button_pos=4; //按扭位置 1左 2右 3上 4下
var stop_time=4000; //图片停留时间(1000为1秒钟)
var show_text=0; //是否显示文字标签 1显示 0不显示
var txtcolor="000000"; //文字色
var bgcolor="DDDDDD"; //背景色
var swf_height=show_text==1?pic_height+20:pic_height;
pics="../../upfile/myfile/xadv/fbot-e1.jpg|../../upfile/myfile/xadv/fbot-e4.jpg|../../upfile/myfile/xadv/fbot-e5.jpg";
links="||";
texts="||";
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cabversion=6,0,0,0" width="'+ pic_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="../../inc/home/playp02.swf">');
document.write('<param name="quality" value="high"><param name="wmode" value="opaque">');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&pic_width='+pic_width+'&pic_height='+pic_height+'&show_text='+show_text+'&txtcolor='+txtcolor+'&bgcolor='+bgcolor+'&button_pos='+button_pos+'&stop_time='+stop_time+'">');
document.write('<embed src="../../inc/home/playp02.swf" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&pic_width='+pic_width+'&pic_height='+pic_height+'&show_text='+show_text+'&txtcolor='+txtcolor+'&bgcolor='+bgcolor+'&button_pos='+button_pos+'&stop_time='+stop_time+'" quality="high" width="'+ pic_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
document.write('</object>');
</script>
<hr>
<iframe id="baiduframe" marginwidth="0" marginheight="0" scrolling="No"
  framespacing="0" vspace="0" hspace="0" frameborder="0" width="230" height="24" 
  src="http://unstat.baidu.com/bdun.bsc?tn=yuyabing_pg&amp;cv=0&amp;cid=1134504&amp;csid=102&amp;bgcr=ffffff&amp;ftcr=000000&amp;urlcr=0000ff&amp;tbsz=110"> </iframe>
<hr />
<table>
  <form method="get" action="http://www.google.cn/search" target="google_window">
    <tr>
      <td><a href="http://www.google.com/"><img src="../../pfile/pimg/google_25wht.gif" width="75" height="32" /></img></a></td>
      <td><input type="text" name="q" size="15" maxlength="255" value="" id="sbi" /></td>
      <td align="left"><input type="submit" name="sa" value="搜索" id="sbb" /></td>
    </tr>
    <input type="hidden" name="client" value="pub-6216376126641226" />
    <input type="hidden" name="forid" value="1" />
    <input type="hidden" name="prog" value="aff" />
    <input type="hidden" name="ie" value="GB2312" />
    <input type="hidden" name="oe" value="GB2312" />
    <input type="hidden" name="cof" value="GALT:#008000;GL:1;DIV:#336699;VLC:663399;AH:center;BGC:FFFFFF;LBGC:336699;ALC:0000FF;LC:0000FF;T:000000;GFNT:0000FF;GIMP:0000FF;FORID:1" />
    <input type="hidden" name="hl" value="zh-CN" />
  </form>
</table>
<hr />
<iframe src="http://www.tianqi123.com/small_page/chengshi_413.html?c0=red&c1=000000&bg=F4F4F4&w=580&h=18&text=yes" width=580 height=18 marginwidth=0 marginheight=0 hspace=0 vspace=2 frameborder=0 scrolling=no align=center id=url></iframe>
<hr />
<SCRIPT language=JavaScript src="http://plus.dg.gov.cn/date8.js"></SCRIPT>
<SCRIPT language=JavaScript src="http://plus.dg.gov.cn/js/dgtq8.js"></SCRIPT>

<hr />
<p>1. id: --- N<br />
  sw01_id1<br />
  sw01_id2</p>
<p>2. css: --- 2N(2) <br />
  sw01_css1a <br />
  sw01_css1b <br />
  sw01_css2a <br />
  sw01_css2b </p>
<p>3. pic: --- 2N(2)<br />
  sw01_pic1a.jpg<br />
  sw01_pic1b.jpg<br />
  sw01_pic2a.jpg<br />
  sw01_pic2b.jpg</p>
<p>5. box: --- N<br />
  sw01_box1<br />
  sw01_box2</p>
<p>6. link: --- 1+N<br />
  sw01_link<br />
  sw01_lnk1<br />
  sw01_lnk2</p>
</body>
</html>
