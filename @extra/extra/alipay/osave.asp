
<!--#include file="alipay/AliPay_Class.asp"-->

<%


ID = Request("ID")
yAct = Request("yAct")


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../pimg/style.css" rel="stylesheet" type="text/css">
<title>网上订购- 东莞汽车旅游中心</title>
<body>
<%


Dim AliPay, AliPayUrl
Set AliPay = New Qlwz_AliPay
AliPay.Subject       = PName    '商品名称
AliPay.Body          = "Order No:" '&OrdNO    '商品描述
AliPay.Price         = "0.01" 'nFee '"0.01"        'price商品单价  0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
AliPay.Quantity      = "1"           '商品数量
'**************以下部分可以在类中设置**************
AliPay.Out_Trade_No  = "Order001" 'OrdNO '按时间获取的订单号
AliPay.Discount      = "0"             '商品折扣
AliPay.Paymethod     = "directPay"     '赋值:bankPay(网银);cartoon(卡通); directPay(余额)
AliPayUrl = AliPay.GetUrl()
Set AliPay = Nothing
itemUrl = AliPayUrl
Response.Write "aa"
Response.end()
'Response.Redirect AliPayUrl   '跳转到支付页

%>
<div class="line03">&nbsp;</div>
<table width="700" border="0" align="center" cellpadding="1" cellspacing="1" style="border:1px #DFE1D4 solid;">
  <tr>
    <td height="26" valign="middle" background="../pimg/lllll_01.jpg"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td width="200" align="left" class="sub04A">网上订购</td>
        <td align="center">&nbsp;</td>
        <td width="120" align="left">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="480" valign="top" style="padding:0px;" class="SysCont">
	
	
      <table width="640" border="0" align="center">
        <form action="?" method="post" name="fm01" id="fm01">
          <tr>
            <td><fieldset style="width:640px; padding:8px; margin:1px 18px;">
              <legend style="font-size:14px;font-weight:bold; color:#C60">定单信息</legend>
              <div style="width:630px; padding:5px;">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#f4f4f4">
                  <tr>
                    <td colspan="2" align="center" valign="middle" bgcolor="#FFFFFF" class="fntF00 " style="line-height:150%; font-size:14px">定单生成成功！感谢您的预定！<br />
                      我们的工作人员会在24小时内与您联系，确认订单并沟通付款事宜。</td>
                  </tr>
                  <tr>
                    <td align="right" valign="middle" bgcolor="#FFFFFF">线 路 名：</td>
                    <td valign="middle" bgcolor="#FFFFFF"><%=PName%></td>
                  </tr>
                  <tr>
                    <td align="right" valign="middle" bgcolor="#FFFFFF">出发日期：</td>
                    <td valign="middle" bgcolor="#FFFFFF"><%=SDate%></td>
                  </tr>
                  <tr>
                    <td align="right" valign="middle" bgcolor="#FFFFFF">定 单 号：</td>
                    <td valign="middle" bgcolor="#FFFFFF"><%=OrdNO%> (<span class="fntF00">请记录此定单号</span>)</td>
                  </tr>
                  <tr>
                    <td width="20%" align="right" valign="middle" bgcolor="#FFFFFF">总 金 额：</td>
                    <td valign="middle" bgcolor="#FFFFFF"><%=nFee%> 元</td>
                  </tr>
                </table>
              </div>
            </fieldset>
              <div class="line10">&nbsp;</div>
              <fieldset style=" width:640px; padding:8px; margin:1px 18px;">
                <legend style="font-size:14px;font-weight:bold; color:#C60">支付方式</legend>
                <div style="width:630px; padding:5px;">
                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
                    <tr bgcolor="#FFFFFF">
                      <td align="left" valign="top" style="line-height:180%; font-size:15px;"> 支付宝帐号：<br />
                        <span class="fntF00">mtc616@yeah.net</span></td>
                      <td width="200" align="center"><a href="<%=itemUrl%>" target="_blank"><img src="alipay/alipay_bwrx.gif" border="0" /></a></td>
                    </tr>
                  </table>
                </div>
                <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" xbgcolor="#E0E0E0">
                  <tr>
                    <td width="50%" align="center" valign="middle" bgcolor="#FFFFFF">&nbsp;</td>
                    <td colspan="4" align="center" valign="middle" bgcolor="#FFFFFF"><input type="button" name="button3" id="button3" value="返回线路" /></td>
                  </tr>
                </table>
              </fieldset>
              <div class="line10">&nbsp;</div></td>
          </tr>
          <input name="Act" type="hidden" id="Act" value="Insert" />
        </form>
      </table>
      <script language="JavaScript" type="text/javascript">

function caFee(e)
{
  var fm = document.fm01; // PAdult,PChild,PBaby,nFee,
  var ev = e.value;
  if ( isNaN(ev) || ev.indexOf(".")>=0 ){
    e.value = 0;
	alert("请填写数字！");
  }
  if((fm.nAdult.value*1+fm.nChild.value*1)>fm.NTotal.value*1){
    e.value = 0;
	alert("成人+小孩,最多可报"+fm.NTotal.value+"人!");
  }
  //fm.nFee.value = fm.PAdult.value*fm.nAdult.value + fm.PChild.value*fm.nChild.value + fm.PBaby.value*fm.nBaby.value;
}


  </script></td>
  </tr>
</table>
<div class="line03">&nbsp;</div>

</body>
</html>