<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>js广告图片播放</title>
</head>
<body>
<div id="TmpID">&nbsp;</div>
<script language="javascript">
function RPStart(xID,xSpeed)
{
  N = eval('RP'+xID+'N');
  var aSrc = eval('RP'+xID+'Src').split(';'); 
  var aUrl = eval('RP'+xID+'Url').split(';'); 
  var oPic = document.getElementById('RP'+xID+'Pic');
  var oLnk = document.getElementById('RP'+xID+'Lnk'); 
  oPic.setAttribute("src",aSrc[N]);
  oLnk.setAttribute("href",aUrl[N]);
  N++;
  if(N==aSrc.length){ eval('RP'+xID+'N=0;'); }
  else              { eval('RP'+xID+'N++;'); }
}
//RPStart('01',1000);
//document.getElementById('TmpID').innerHTML = "<BR>"+aSrc[RP01N]+oLnk+"<BR>";
//setInterval("RPStart('"+xID+"',"+xSpeed+")",1000);
</script>
<a id="RP01Lnk" href="#"><img id="RP01Pic" src="../../img/logo/y_logo.gif" width="180" height="140" border="0" /></a>
<script language="javascript">
var RP01Src = "../../img/logo/temp-qingzhu.jpg;../../img/logo/temp_2008.jpg;../../img/logo/temp_480x200.jpg";
var RP01Url = "01.htm;02.htm;03.htm";
var RP01N = 0;
setTimeout("RPStart('01')",300);
setInterval("RPStart('01')",2000);
</script>
<br />
<a id="RP02Lnk" href="#"><img src="../../img/logo/y_logo.gif" name="RP02Pic" width="55" height="86" border="0" id="RP02Pic" /></a>
<script language="javascript">
var RP02Src = "../../img/logo/wj_chacha.gif;../../img/logo/wj_jinjin.gif;../../img/logo/net110gt.gif;../../upfile/myftp/test/clock03.swf";
var RP02Url = "21.htm;22.htm;23.htm";
var RP02N = 0;
setTimeout("RPStart('02')",23);
setInterval("RPStart('02')",1500);
</script>
<br />
<a id="TmpLink" href="#">Peace Test Text Box</a>
<script language="javascript">
function TmpStart()
{
  var tLink = document.getElementById('TmpLink');
  tLink.innerHTML = '<embed src="../../upfile/myftp/test/clock03.swf" quality="high" wmode="opaque" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="150" height="150"></embed>';
}
setTimeout("TmpStart()",800);
</script>
</body>
</html>
