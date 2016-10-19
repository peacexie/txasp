<!--#include file="inc_func.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Oracle Test</title>
<style type="text/css">
body,td,th {
	font-size: 12px;
}
td {
  background-color:#FFF;
  padding:1px;
}
</style>
</head>
<body>
<%

yAct = Request("yAct")
'Response.Write Enc_Peace("padm")

If yAct="SetUser" Then
  Session("UsrID") = "peace(DB)"
  Response.Write "User"
End If
If yAct<>"" AND Session("UsrID")&""<>"" Then  '
  Session("PeaceDB") = Enc_Peace(yAct)
  Response.Write "OK"
End If
If Session("PeaceDB")&""<>PassID Then
%>
<table width="240" border="1" align="center">
<form id="form1" name="form1" method="post" action="?">
  <tr>
    <td><a href="?yAct=SetUser">Login</a></td>
    <td><input name="yAct" type="password" id="yAct" value="112233" size="24" maxlength="24" />
    </td>
    <td><input type="submit" name="button" id="button" value="提交" /></td>
  </tr>
  </form>
</table>
<%
Else
%>
<table border="1" align="center" bgcolor="#F0F0F0">
  <form name="form2" method="post" action="">
  <tr>
    <td align="right">Test</td>
    <td><input name="xTab" type="text" id="xTab" value="<%=xTab%>" size="32" maxlength="24" /></td>
    <td><input type="submit" name="button2" id="button2" value="提交" /></td>
  </tr>
<%
xTab = RequestS("xTab","C",64)
If xTab<>"" Then
%>
  <tr>
    <td>OK:</td>
    <td><%=xTab%></td>
    <td><%=RequestSafe(rs_Count(conn,xTab),"N",0)%></td>
  </tr>
<%End If%>
</form>  
</table>
<div style="line-height:12px">&nbsp;</div>
<table width="580"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#F0F0F0">
  <tr>
    <td width="20%" align="center">&nbsp;</td>
    <td width="20%" align="center">&nbsp;</td>
    <td width="20%" align="center">&nbsp;</td>
    <td width="20%" align="center">&nbsp;</td>
    <td width="20%" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><a href="data-p1.asp">Data-Type1</a></td>
    <td align="center"><a href="data-p2.asp">Data-Type2</a></td>
    <td align="center"><a href="data_rnum.asp">Data-RowNum</a></td>
    <td align="center"><a href="data_rnum2.asp">Data-RowNum2</a></td>
    <td align="center"><a href="data_rnum3.asp">Data-RowNum3</a></td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center"><a href="sys_max.asp">Sys-Max</a></td>
    <td align="center"><a href="sys_max2.asp">Sys-Max</a></td>
    <td align="center"><a href="ora_max.asp">Data-Max</a></td>
  </tr>
  <tr>
    <td align="center"><a href="cust_01.asp">CST_Customer1</a></td>
    <td align="center"><a href="cust_02.asp">CST_Customer2</a></td>
    <td align="center"><a href="cust_03.asp?xTab=CST_CUSTOMER">CST_03A</a></td>
    <td align="center"><a href="cust_03.asp?xTab=SYS_STAFF">CST_03B</a></td>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><a href="ord_01.asp?xTab=SALE_ORDER">Sale_Order1</a></td>
    <td align="center"><a href="ord_01.asp?xTab=SALE_GROUP_ORDER">Sale_OrdSubs</a></td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
</table>
<div style="line-height:12px">&nbsp;</div>
<table width="580"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#F0F0F0">
  <tr>
    <td align="right">&nbsp;**</td>
    <td>DOMAIN_PRODUCT_TYPE</td>
    <td align="left">产品类型树</td>
    <td align="right">***</td>
    <td align="center"><a href="ora_stru.asp?xTab=DOMAIN_PRODUCT_TYPE">Stru</a></td>
    <td align="center"><a href="ora_data.asp?xTab=DOMAIN_PRODUCT_TYPE">Data</a></td>
  </tr>
</table>
<!--#include file="inc_tabs.htm"-->
<%
  'Response.Redirect "ora_tab.asp"
End If

%>
</body>
</html>
