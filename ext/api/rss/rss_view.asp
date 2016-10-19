<!--#include file="rss_config.asp"-->
<%
KeyID = RequestS("KeyID",3,48)
TypID = Request("TypID") 
nID = fChkRID(TypID)


If nID>=0 Then
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM "&aTabID(nID)&" WHERE KeyID='"&KeyID&"'",conn,1,1 
  if NOT rs.eof then 
  Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = Show_Text(rs("InfSubj"))
InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
SetRead = rs("SetRead")+1
LogATime = rs("LogATime")
InfCont = Show_Html(rs("InfCont"))
ImgName = get_1Img(KeyID,rs("ImgName")) 
If ImgName<>"" Then '/img/tool/no_pic
  ImgName = "<center><a href='"&ImgName&"' target='_blank'><img src='"&ImgName&"' width='640' xheight='720' onload='javascript:setImgSize(this);' border=0 /></a></center><br>"
  InfCont = ImgName&InfCont
End If
  rs.MoveNext
  Loop
  Else
KeyID = ""
  End If
rs.Close()
SET rs=Nothing
End If 

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=InfSubj%></title>
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
<!--
body,td,th {
	font-size: 13px;
}
body {
	margin-left: 24px;
	margin-right: 24px;
	margin-top: 18px;
	margin-bottom: 24px;
}
-->
</style>
</head>
<body>

<%
If KeyID="" OR nID<0 Then 
  Response.Write "<h3 align='center'>错误！！！</h3>"&vbcrlf&"</body>"&vbcrlf&"</html>" 
  Response.End() 
Else
  Call rs_DoSql(conn,"UPDATE InfoNews SET SetRead=SetRead+1 WHERE KeyID='"&KeyID&"'")
End If
%>

<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1">
  <tr>
    <td height="30" colspan="2" style="font-size:14px; border:1px solid #CCC; background-color:#FFC;">&nbsp;(RSS订阅) --- <%=Config_Name%> --- <%=aTitle(nID)%> --- 详情</td>
  </tr>
  <tr>
    <th height="40" colspan="2" style="font-size:16px; color:#36C"><strong><%=InfSubj%></strong></th>
  </tr>
  <tr>
    <td nowrap bgcolor="#F0F0F0">Publish: <%=LogATime%></td>
    <td nowrap bgcolor="#F0F0F0">Read: <%=SetRead%></td>
  </tr>
  <tr>
    <th height="24" colspan="2"><hr color="#3366CC"></th>
  </tr>
  <tr>
    <td colspan="2" style="font-size:14px; line-height:150%;"><%=InfCont%></td>
  </tr>
  <tr>
    <th height="24" colspan="2"><hr color="#3366CC"></th>
  </tr>
  <%If aTabID(nID)="InfoNews" Then%>
  <tr>
    <td bgcolor="#F0F0F0">From: <%=InfFrom%><br>
      Keywords：<%=InfKey%></td>
    <td width="30%" nowrap bgcolor="#F0F0F0">(<%=LogAddIP%>)</td>
  </tr>
  <%Else%>
  <tr>
    <td bgcolor="#F0F0F0">Speci: <%=InfSpeci%></td>
    <td width="30%" nowrap bgcolor="#F0F0F0">(<%=LogAddIP%>)</td>
  </tr>
  <%End If%>
</table>
</body>
</html>
