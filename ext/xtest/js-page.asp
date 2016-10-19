<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>jsPage-Demo</title>
<link href="../../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
.p {
	width:90px;
	display:inline-table;
	float:left;
	padding:1px 8px;
	margin:0px;
}
</style>
</head>
<body>
<div style="width:640px; margin:30px 0px 10px 80px; padding:12px; background-color:#F0F0F0">
  <%
For i=1 To 36
  j = i
  Response.Write vbcrlf&"<div class='p'><a href='?p="&j&"'>All["&j&"]页</a></div>"
Next
p = Request("p")
If p="" Then p=8
%>
</div>
<div style="width:640px; margin:10px 0px 20px 80px; padding:12px; background-color:#F0F0F0">
  <div id="pageData"> 信息开始 ......
    第1页内容 ......<br />
    ...... <br />
    <%For i=2 To p%>
    <div style="page-break-after: always"><span style="display: none">&nbsp;</span></div>
    第<%=i%>页内容 --- <br />
    美俄紧锣密鼓准备“换谍”; 伊戈尔·苏佳金已被转移至莫斯科的列夫特沃监狱。 <br />
    被关押的俄罗斯核武专家伊戈尔·苏佳金。 <br />
    （资料图片） 伊戈尔·苏佳金的律师（右）和家人7日接受采访时透露了“换谍”细节。 <br />
    深圳特区报讯据外国媒体报道，美俄两国 ... <%=i%>
    <%Next%>
    <br />
    ...... <br />
    信息结束了！！！ </div>
  <div style="clear:both;"></div>
  <div id="pageOne" style="display: none">[pageOne]</div>
  <div style="clear:both;"></div>
  <TABLE id="pageBox" border="0" align=center cellSpacing=0 style="MARGIN: 12px auto 0px; display: none">
    <TR>
      <TD id="pageBar" class="pageCell"> [pageBar copy@peace.xie.ys 2010-07-08]</TD>
    </TR>
  </TABLE>
  <script src="../../inc/home/jsPager.js" type="text/javascript"></script>
</div>
<hr />
<p>&lt;&gt; <div style="width:3px; float:left;"></div> <SPAN style="width:3px;"></SPAN> </p>
</body>
</html>
