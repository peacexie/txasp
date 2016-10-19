<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE>外部js - for www.dg.gd.cn</TITLE>
<META http-equiv=Pragma content=no-cache>
<style type="text/css">
<!--
body {
	margin: 1px;
}
body, td, th {
	font-size: 12px;
}
-->
</style>
</HEAD>
<BODY bgcolor="#F0F0F0">
<!--#include file="_config.asp"-->
<!--#include file="../../sadm/func1/funcfile.asp"-->
刷新:OK …… 文化周末 …… 外部js - for www.dg.gd.cn
<hr>
<%

'第"&&"期 <a href='http://whzm.dg.gd.cn/page/pview.asp?KeyID="&&"' target='_blank'>"&&"</a><br />
s = ""

mMonth = rs_Val(conn,"SELECT TOP 1 KeyID AS InfSubj FROM OrdPara WHERE KeyMod='MName' AND InfCont<>'Hidden' ORDER BY KeyID DESC ","InfSubj")&""
Response.Write "<br> Max Pos ============= "&mMonth
'YYYYMM = Left(mMonth,4)&"年"&Right(mMonth,2)&"月"
'MinID = rs_Val(conn,"SELECT Min(InfPrice) AS NowID FROM [InfoPics] WHERE KeyMod='PicT124' AND InfPric2="&mMonth&"","NowID")
'MaxID = rs_Val(conn,"SELECT Max(InfPrice) AS NowID FROM [InfoPics] WHERE KeyMod='PicT124' AND InfPric2="&mMonth&"","NowID")

sql = "SELECT TOP 2 * FROM [InfoPics] WHERE KeyMod='PicT124' AND InfPric2<="&mMonth&" ORDER BY InfPrice DESC,LogATime DESC " 
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
 KeyID = rs("KeyID")
 InfSubj = rs("InfSubj") 
 InfSubj = InfSubj&"——"&rs("InfSubj2")
 InfSubj = Show_SLen(InfSubj,24)
 InfSubj = Replace(InfSubj,"'","‘")
 InfSubj = Replace(InfSubj,"""","”")
 InfSubj = Replace(InfSubj,"————","——")
 InfSubj = Replace(InfSubj,vbcrlf,"")
 InfSubj = Replace(InfSubj,vbcr,"")
 InfSubj = Replace(InfSubj,vblf,"")
 InfPrice = rs("InfPrice")
 LogATime = rs("KeyCode") 
 istr = "<div style='line-height:120%; padding:1px 1px 1px 1px; '>第"&InfPrice&"期 "
 istr = istr& " <a href='http://whzm.dg.gd.cn/page/pview.asp?KeyID="&KeyID&"' target='_blank' title='"&LogATime&"'>"&InfSubj&"</a></div>"
 s = s&vbcrlf&"document.write("""&istr&""");"
rs.MoveNext
Loop
rs.Close()

Response.Write "<br>"&s
Call File_Add2("/script/sys/dg.gd.js",s,"gb2312")

Set rs=Nothing

%>
<hr>
前台调用：<br>
&lt;script language=&quot;javascript&quot; src=&quot;http://whzm.dg.gd.cn/script/sys/dg.gd.js&quot; charset='gb2312'&gt;&lt;/script&gt; <br>
<script language="javascript" src="http://whzm.dg.gd.cn/script/sys/dg.gd.js?a=<%=Timer()%>" charset='gb2312'></script>
<br />
后台刷新：<br />
&lt;IFRAME name='fOutjs' src=&quot;http://whzm.dg.gd.cn/ext/out_js.asp&quot; frameBorder=0 width=&quot;100%&quot; scrolling=yes height=&quot;15&quot;&gt;
</IFRAME>
</BODY>
</HTML>
