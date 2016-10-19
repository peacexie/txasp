<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="func_const.asp"-->
<!--#include file="cch_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Demo - Cache</title>
</head>
<body>


<%

'Response.Write "<br>fileExist('/index.asp') : "&fileExist("/index.asp")
'Response.Write "<br>fileRead('/index.asp') : [[["&fileRead("/index.asp","utf-8") &"]]]" 
'Response.Write "<br>urlPath('/index.asp') : "&urlPath("/index.asp")
'Response.Write "<br>urlRead('/index.asp') : [[["&urlRead("/index.asp") &"]]]"
'Response.Write "<br>urlSave('/index.asp') : [[["&urlSave("/index.asp","/upfile/sys/cache/00home.htm") &"]]]"  

Response.Write "<hr>"
'Response.End()

%>
<hr>
<%

'--------------------------------使用方法
dim sCont,myCache,myName,myTime
myName = "TestCont"
myTime = 10 ' 时间-分钟(10-60)24*6=144
set myCache = new clsCache
  myCache.name = myName '定义缓存名称
  If Request("act")="Clear" Then
    myCache.clean()
  End If
  if myCache.valid then '如果缓存有效
    sCont = myCache.value '读取缓存内容
    a = "您看到的是缓存"
  else
    'sCont = "看到这个，您就看到了先锋缓存类！" '大量内容，可以是非常耗时大量数据库查询记录集
	sCont = urlRead(urlPath(Config_Path&"page/index.asp")) 
    myCache.add sCont,dateadd("n",myTime,now) '将内容赋值给缓存，并设置缓存有效期是当前时间＋N分钟
    a = "将内容写入缓存"
  end if 
  response.write sCont & "<BR>"& a &"<BR>"
  response.write myCache.Version
set myCache = nothing '释放对象


'--------------------------------使用说明
%>
<hr />
<a href="?">Normal</a> | <a href="?act=Clear">Clear()</a> 
<TABLE WIDTH="500" BORDER="0" CELLPADDING="5" CELLSPACING="0" CLASS="p9">
  <TR>
    <TD HEIGHT="10" COLSPAN="2"></TD>
  </TR>
  <TR>
    <TD COLSPAN="2"><STRONG>clsCache 公共属性</STRONG></TD>
  </TR>
  <TR>
    <TD WIDTH="30">&nbsp;</TD>
    <TD><TABLE WIDTH="100%" BORDER="1" CELLPADDING="0" CELLSPACING="0" BORDERCOLOR="#999999" CLASS="stable">
        <TR VALIGN="middle">
          <TD WIDTH="210" HEIGHT="20"><IMG SRC="1.GIF" WIDTH="16" HEIGHT="16"> valid </TD>
          <TD>&nbsp;返回是否有效。true表示有效,false表示无效。只读。</TD>
        </TR>
        <TR VALIGN="middle">
          <TD WIDTH="210" HEIGHT="20"><IMG SRC="1.GIF" WIDTH="16" HEIGHT="16"> Version</TD>
          <TD>&nbsp;获取类的版本信息。只读。</TD>
        </TR>
        <TR VALIGN="middle">
          <TD HEIGHT="20"><IMG SRC="1.GIF" WIDTH="16" HEIGHT="16"> value </TD>
          <TD>&nbsp;返回缓存内容。只读。</TD>
        </TR>
        <TR VALIGN="middle">
          <TD HEIGHT="20"><IMG SRC="1.GIF" WIDTH="16" HEIGHT="16"> name</TD>
          <TD>&nbsp;设置缓存名称，写入。</TD>
        </TR>
        <TR VALIGN="middle">
          <TD HEIGHT="20"><IMG SRC="1.GIF" WIDTH="16" HEIGHT="16"> expire</TD>
          <TD>&nbsp;设置缓存过期时间，写入。</TD>
        </TR>
      </TABLE></TD>
  </TR>
  <TR>
    <TD HEIGHT="10" COLSPAN="2"></TD>
  </TR>
  <TR>
    <TD COLSPAN="2"><STRONG>clsCache 公共方法</STRONG></TD>
  </TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD><TABLE WIDTH="100%" BORDER="1" CELLPADDING="0" CELLSPACING="0" BORDERCOLOR="#999999" CLASS="stable">
        <TR VALIGN="middle">
          <TD WIDTH="210" HEIGHT="20"><IMG SRC="2.GIF" WIDTH="16" HEIGHT="16"> add(Cache,ExpireTime)</TD>
          <TD>&nbsp;对缓存赋值（缓存内容,过期时间）</TD>
        </TR>
        <TR VALIGN="middle">
          <TD HEIGHT="20"><IMG SRC="2.GIF" WIDTH="16" HEIGHT="16"> clean()</TD>
          <TD>&nbsp;清空缓存</TD>
        </TR>
        <TR VALIGN="middle">
          <TD HEIGHT="20"><IMG SRC="2.GIF" WIDTH="16" HEIGHT="16"> verify(cchCont2)</TD>
          <TD>&nbsp;比较缓存值是否相同——返回是或否。</TD>
        </TR>
      </TABLE></TD>
  </TR>
  <TR>
    <TD HEIGHT="5" COLSPAN="2">&nbsp;</TD>
  </TR>
</TABLE>
</body>
</html>
