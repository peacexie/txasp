<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Function getPara(xPara) ' Save,Edit,UBB,show text in html
   Dim xText  
   xText = Request(xPara)
   xText = Replace(xText, "<", "&lt;")
   xText = Replace(xText, ">", "&gt;")    
   xText = Replace(xText, " ", "&nbsp;")
   getPara = xText
End Function 
sUrl = getPara("sUrl") 
Subj = getPara("Subj") 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<title><%=Subj%></title>
<noscript><iframe src='*.PeacePage'></iframe></noscript>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../../inc/home/jskdis.js" type="text/javascript"></script>
</head>
<body>
<table width="640" border="0" align="center" cellspacing="8">
  <tr>
    <td align="center"><img src="<%=sUrl%>" alt="<%=Subj%>" width="480" height="360" border="0" id="ImgID" onload="javascript:setImgSize(this);setTimeout('setLink()',800);"></td>
  </tr>
  <tr>
    <td align="center" id="LinkID"><%=Subj%></td>
  </tr>
</table>
<script type="text/javascript">
function setLink(){
  var img = getElmID("ImgID");
  var w = img.width; 
  var h = img.height;
  if((w>480)||(h>360)){
	 var lnk = getElmID("LinkID");
	 lnk.innerHTML = "<a href='"+img.src+"' style='text-decoration:none'>"+lnk.innerHTML+"</a>";
  }
}
//setTimeout("setLink()",1200);
</script>
</body>
</html>
