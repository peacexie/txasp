
<!--#include file="config.asp"-->
<!--#include file="../../pfile/lang/vmemb.asp"-->
<%



ID = RequestS("ID",3,48) 

SET rs=Server.CreateObject("Adodb.Recordset") 

rs.Open "SELECT * FROM OrdInfo WHERE KeyID='"&ID&"' ",conn,1,1 
  if NOT rs.eof then 
  Do While Not rs.EOF

KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))
LogAddIP = rs("LogAddIP") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")

LnkName   = rs("LnkName")
LnkSex    = rs("LnkSex")
LnkAddr   = rs("LnkAddr")
LnkPost   = rs("LnkPost")
LnkMobile = rs("LnkMobile")
LnkTel    = rs("LnkTel")
LnkEmail  = rs("LnkEmail")
 
InfNum   = rs("InfNum")
InfPay   = rs("InfPay")
InfSend  = rs("InfSend")
InfTime  = rs("InfTime")

SendType = rs_Val("","SELECT InfCont FROM OrdPara WHERE KeyID='"&InfSend&"'")
SendTime = rs_Val("","SELECT InfCont FROM OrdPara WHERE KeyID='"&InfTime&"'")
PayType  = rs_Val("","SELECT InfSubj FROM OrdPara WHERE KeyID='"&InfPay&"'")
 
SetCheck  = rs("SetCheck")
SetPay    = rs("SetPay")
SetSend   = rs("SetSend")
SetState  = rs("SetState")

  rs.MoveNext
  Loop
  end if 
rs.Close()

If KeyID = "" Then 
  Response.End() '.Redirect("adm_list.asp")
End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=KeyCode%> --- 定单查看</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.OrdFLeg {
	padding:2px 5px 2px 5px;
	font-weight:bold;
}
.OrdFSet {
	padding:3px 5px 5px 5px;
	margin:2px 5px;
}
</style>
</head>
<body>
<table width="680" border="0" align="center" cellpadding="0" cellspacing="0">

    <tr>
      <td height="500" colspan="4" align="left" valign="top" style="BORDER:#E0E0E0 1px solid; padding:5px;"><!--Item Start-->
        
<fieldset class="OrdFSet">
          <legend class="OrdFLeg">订单信息 [<%=KeyCode%>]</legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center">定单标题</td>
              <td colspan="3"><%=InfSubj%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">定单编号</td>
              <td width="30%"><%=KeyCode%></td>
              <td width="20%" align="center">合计金额</td>
              <td width="30%"><%=InfNum%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">订购时间</td>
              <td><%=LogATime%></td>
              <td align="center">订购ＩＤ</td>
              <td><%=LogAUser%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">配送方式</td>
              <td colspan="3"><%=SendType%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">时间要求</td>
              <td colspan="3"><%=SendTime%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">支付方式</td>
              <td colspan="3"><%=PayType%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">附加留言<br />
                (120字内)</td>
              <td colspan="3"><%=InfCont%></td>
            </tr>
          </table>
        </fieldset>
        
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg">货款信息</legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td align="center" bgcolor="#FFFFF0">NO.</td>
              <td width="40%" bgcolor="#FFFFF0">商品名称</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">价格</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">数量</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">小计</td>
            </tr>
            <%
		 
sql = "SELECT * FROM [OrdItem] "
sql = sql & " WHERE KeyCode='"&ID&"' "
sql = sql & " ORDER BY KeyID "
'Response.Write sql
aNum = 0
aSum = 0
aStr = ""

rs.Open sql,conn,1,1 
i = 0
Do WHILE NOT rs.EOF
 i = i +1
if i mod 2 = 1 then
   col = "F8F8F8"
else
   col = "FFFFF8"
end if

InfProID = rs("InfProID")
iSubj = rs_Val("","SELECT InfSubj FROM InfoPics WHERE KeyID='"&InfProID&"'")
InfPrice = rs("InfPrice")
InfCount = rs("InfCount")
InfSum = rs("InfPrice")*rs("InfCount")
aSum = aSum+InfSum
aNum = aNum+InfCount
aStr = aStr&InfProID&";"
%>
            <tr bgcolor="<%=col%>">
              <td align="center" nowrap="nowrap"><%=i%></td>
              <td><%=iSubj%></td>
              <td nowrap="nowrap"><%=InfPrice%></td>
              <td nowrap="nowrap"><%=InfCount%></td>
              <td nowrap="nowrap"><%=InfSum%></td>
            </tr>
            <% 
rs.MoveNext 
Loop 
rs.Close()
%>
            <tr bgcolor="f0f0f0">
              <td align="center" nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;合计&nbsp;</td>
              <td bgcolor="#FFFFF0">&nbsp;</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><%=aNum%></td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><%=aSum%></td>
            </tr>
          </table>
        </fieldset>
       
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg">送货地址 和 联系方式</legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center">收货人</td>
              <td width="30%"><%=LnkName%></td>
              <td width="20%" align="center">称 谓</td>
              <td width="30%"><%=LnkSex%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">地 址</td>
              <td colspan="3"><%=LnkAddr%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">邮政编码</td>
              <td><%=LnkPost%></td>
              <td align="center">手 机</td>
              <td><%=LnkMobile%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center">电 话</td>
              <td><%=LnkTel%></td>
              <td align="center">E-mail</td>
              <td><%=LnkEmail%></td>
            </tr>
          </table>
        </fieldset>
        
        
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg">定单操作</legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="25%" align="center"><a href="javascript:window.print()">[打印定单]</a></td>
              <td width="25%" align="center">订购时间</td>
              <td width="25%" align="center"><%=LogATime%></td>
              <td width="25%" align="center">
                <a href="javascript:window.close()">[关闭窗口]</a>
              </td>
            </tr>
          </table>
        </fieldset>
        <!--Item End--></td>
    </tr>
</table>
<%
SET rs=Nothing 
%>
</body>
</html>
