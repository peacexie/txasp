<%
  'If eflg Then
	'Response.Redirect "../page/message.asp?mType=Error&msg="&emsg&"&goPage=goRef"
	'Response.End()   
  'End If 
mType = Request("mType") 'Error
If mType="Error" Then
  mSubj = "错误信息"
ElseIf mType="contMode" Then
  mSubj = "内容模式说明"
ElseIf mType="Debug" Then
  mSubj = "接口调试说明"
Else
  mSubj = ""
End If
msg = Request("msg")
goPage = Request("goPage")
If goPage="goRef" Then
  goPage = Request.Servervariables("HTTP_REFERER")
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mSubj%></title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css"/>
</head>
<body>
<%If mType="Error" Then%>
<table width="610" border="1" align="center" style="margin:20px 0px;">
  <tr>
    <td align="left" class="cRed f14px">&nbsp;&nbsp; 错误！可能原因如下：</td>
  </tr>
  <tr>
    <td class="f14px">1. <%=msg%>， 请<a href="<%=goPage%>" class="cBlue">返回</a>！</td>
  </tr>
  <tr>
    <td class="f14px">2. 或者超时，请重新登陆！</td>
  </tr>
</table>
<%End If%>
<%If mType="contMode" Then%>
<table width="610" border="1" align="center" style="margin:5px;">
  <tr>
    <td align="left" class="cRed f14px">&nbsp;&nbsp; 内容模式说明：</td>
  </tr>
  <tr>
    <td class="f14px">1. 兼容模式 --- 推荐！一次只发70个字以内</td>
  </tr>
  <tr>
    <td class="f14px">2. 手动分割 --- 推荐！最多4条信息,每条70个字以内</td>
  </tr>
  <tr>
    <td valign="top" class="f14px">3. [长文本] --- 最多255个字(兼容性差)<br />
      长短信扣费说明: 如果一次提交小于等于70字符 系统会默认为一条短信发出 扣费一条; <br />
      如果大于70字符,系统会默认为长短信处理,此时扣费按65字符扣一条; 所以198字符 是 65*3+3  扣费4条; 对方如果手机支持长短信则收到一条, 如果不支持长短信 则收到四条。 (建议内容不超过180个字内,扣费3条以内)</td>
  </tr>
  <tr>
    <td valign="top">
      <div class="f14px" style="height:256px; overflow-x:hidden;overflow-y:scroll;">
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>关于手机短信长度问题<br />
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>以前一直疑惑于手机短信长度问题。在美国，GSM手机短信（SMS）的长度为160个字符，但是，在中国，中文手机短信的长度是70个字符。我一直在考虑，为什么长度不是80个字符呢？原来总是认为两个英文字符算是一个中文字符。<br />
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>后来，仔细研读关于手机短信的技术文献，得知，一个手机短信的实际长度是140个字节，但是，普通的英文短信都是使用七位字符编码。这样，140 x 8 / 7 = 160个字符，原来如此。<br />
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>这样，140个字节正好编码70个中文字符。但是，如果是中英文混合短信呢？<br />
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>原来，对于手机短信，有三种编码方式，七位，八位，和Unicode UTF-16三种。140个字节可以用来编码160个七位字符，140个八位字符，和70个UTF-16字符。只要短信中间出现中文，则这个短信就要用UTF-16编码，在UTF-16中，英文字符也需要两个字节来编码。所以中英文混合短信也只能有七十个字符，不管是中文还是英文的。<br />
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>如果是纯英文短信呢？根据使用的手机，这个短信可能被编码成七位，八位，或者UTF-16三种方式，对应于160，140，和70字符三种长度。到底使用哪个编码方式应该取决于所使用的手机。美国的手机估计都会用七位方式编码，得到最大的字符长度160个字符，在中国买的中文手机有可能使用UTF-16来编码，这样得到的字符长度最短，只有七十个，跟中文短信长度相同。也有可能使用八位编码，这样一个短信就有140个字符长度。<br />
      &nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>同样型号的手机刷了不同的ROM也可能使用不同的方式编码，这个只有经过试验才知道。我现在还没有仔细试过不同的中文手机来发纯英文短信，所以没有确定的答复。<br />
      <span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span><span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span>by:神游四海 (Tao’s Weblog) http://www.tongtao.com/myblog/index.php?p=190<br />
      <span class="f14px" style="height:256px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;<span class="f14px" style="height:180px; overflow-x:hidden;overflow-y:scroll;">&nbsp;&nbsp;</span></span>(Peace)很多手机，最多可发260个文字，这个数字=65x4，其实是按4条短信扣费。</div>
      </td>
  </tr>
</table>
<%End If%>
<%If mType="Debug" Then%>
<table width="610" border="1" align="center" style="margin:20px 0px;">
  <tr>
    <td align="left" class="cRed f14px">&nbsp;&nbsp; 公告！接口调试说明：</td>
  </tr>
  <tr>
    <td class="f14px">1. 接口目前正在调试，可能是对系统作一些小的调整升级等，请稍等一会再试！ </td>
  </tr>
  <tr>
    <td class="f14px">2. 请<a href="<%=goPage%>" class="cBlue">返回</a>！</td>
  </tr>
</table>
<%End If%>
</body>
</html>
