
<!--#include file="alipay/AliPay_Class.asp"-->

<%


ID = Request("ID")
yAct = Request("yAct")


%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>定单详情</title>
<link href="../../inc/mem_inc/mem_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.Inp23 {
	width:30px;
	height:20px;
	font-size:12px;
	line-height:100%;
	padding:2px;
	margin:0px;
}
</style>
</head>
<body>

<table border="0" width="100%" bgcolor="#3C5146" cellspacing="1" cellpadding="7">
  <tr>
    <td width="100%" bgcolor="#F1FFF0" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E6E6E6">
      <tr>
        <td bgcolor="#FAFAFA"  align="right" width="19%">订单编号：</td>
        <td bgcolor="#FAFAFA"  width="31%"><%=OrdNO%></td>
        <td bgcolor="#FAFAFA"  align="right" width="19%">状态：</td>
        <td bgcolor="#FAFAFA"  width="31%"><%=fState%><span style="color:#F1FFF0"><%=OrdState%></span> [<a href="?ID=<%=ID%>" title="如果最近有支付操作，点此[刷新]查看最新状态。">刷新</a>]</td>
      </tr>
      <tr>
        <td bgcolor="#FAFAFA"  align="right">订购者：</td>
        <td bgcolor="#FAFAFA"  ><%=uLogin%></td>
        <td bgcolor="#FAFAFA"  align="right">联系人及电话：</td>
        <td bgcolor="#FAFAFA"  ><%=OrdHand%></td>
      </tr>
      <tr>
        <td bgcolor="#FAFAFA"  align="right">创建时间：</td>
        <td bgcolor="#FAFAFA"  ><%=OrdAddTM%></td>
        <td bgcolor="#FAFAFA"  align="right">修改时间：</td>
        <td bgcolor="#FAFAFA"  ><%=OrdUpdTM%></td>
      </tr>
      <tr>
        <td bgcolor="#FAFAFA"  align="left" colspan="4">&nbsp;&nbsp;<b>子订单信息：</b> <span class="fntF00"><%=Msg%></span></td>
      </tr>
      <tr>
        <td colspan="4" align="center"  bgcolor="#FAFAFA" style="padding:5px;"><table width="98%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
          <tr height="22"  class="trbg"  style="cursor: default">
            <td align="center" bgcolor="#FFFFFF"><b>子订单编号</b></td>
            <td width="40%" align="center" bgcolor="#FFFFFF"><b>产品</b></td>
            <td align="center" nowrap bgcolor="#FFFFFF">成年价<br>
              成年数</td>
            <td align="center" nowrap bgcolor="#FFFFFF">小孩价<br>
              小孩数</td>
            <td align="center" nowrap bgcolor="#FFFFFF">婴儿价<br>
              婴儿数</td>
            <td align="center" nowrap bgcolor="#FFFFFF"><b>应收</b></td>
            <td align="center" nowrap bgcolor="#FFFFFF"><b>状态</b></td>
            <td align="center" nowrap bgcolor="#FFFFFF"><b>操作</b></td>
          </tr>
          
          <form action="?" name="fm01<%=SubID%>" method="post">
            <tr height="20" subOrderID="284929"state="2" bgcolor="#FFFFFF" style="cursor: default">
              <td align="center" nowrap bgcolor="#FFFFFF"><font color='#BB1010'> <%=SubNumb%></font></td>
              <td bgcolor="#FFFFFF"><font color='#BB1010'><%=pName%>
                <input name="NTotal" type="hidden" id="NTotal" value="<%=NTotal%>">
                <input name="nFee" type="hidden" id="nFee" value="<%=RECEIVABLE%>">
                <input name="yAct" type="hidden" id="yAct" value="Upd">
                <input name="SubID" type="hidden" id="SubID" value="<%=SubID%>">
                <input name="ID" type="hidden" id="ID" value="<%=ID%>">
              </font></td>
              <td align="center" bgcolor="#FFFFFF"><%=P1%>
                <input name="PAdult" type="hidden" id="PAdult" value="<%=P1%>">
                <br>
                <input name="nAdult" type="text" class="Inp23" id="nAdult" value="<%=Q1%>" size="5" maxlength="4" onBlur="caFee(this,'<%=SubID%>')"></td>
              <td align="center" bgcolor="#FFFFFF"><%=P2%>
                <input name="PChild" type="hidden" id="PChild" value="<%=P2%>">
                <br>
                <input name="nChild" type="text" class="Inp23" id="nChild" value="<%=Q2%>" size="5" maxlength="4" onBlur="caFee(this,'<%=SubID%>')"></td>
              <td align="center" bgcolor="#FFFFFF"><%=P3%>
                <input name="PBaby" type="hidden" id="PBaby" value="<%=P3%>">
                <br>
                <input name="nBaby" type="text" class="Inp23" id="nBaby" value="<%=Q3%>" size="5" maxlength="4" onBlur="caFee(this,'<%=SubID%>')"></td>
              <td align="center" bgcolor="#FFFFFF"><span id="M1_<%=SubID%>"><%=RECEIVABLE%></span></td>
              <td align="center" bgcolor="#FFFFFF"><%=sState%></td>
              <TD align="center" bgcolor="#FFFFFF"><input type="submit" name="button" id="button" value="修改" <%=fDis%>></TD>
            </tr>
          </form>

          <input type="hidden" name="hasSubOrder" value="true" >
          <tr>
            <td align="center" bgcolor="#FFFFFF"><b>合计</b></td>
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF"><span id="M9_All"><%=All01%></span></td>
            <td colspan="2" align="center" bgcolor="#FFFFFF"><a href="<%=itemUrl%>" target="_blank"></a></td>
          </tr>
        </table>
          <%

	
Dim AliPay, AliPayUrl
Set AliPay = New Qlwz_AliPay
AliPay.Subject       = "Test商品名称" 'pName    '商品名称
AliPay.Body          = "Order No:"&OrdNO    '商品描述
AliPay.Price         = "0.01" 'All01 '"0.01"        'price商品单价  0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
AliPay.Quantity      = "1"           '商品数量
'**************以下部分可以在类中设置**************
AliPay.Out_Trade_No  = "Order001" 'OrdNO '按时间获取的订单号
AliPay.Discount      = "0"             '商品折扣
AliPay.Paymethod     = "directPay"     '赋值:bankPay(网银);cartoon(卡通); directPay(余额)
AliPayUrl = AliPay.GetUrl()
Set AliPay = Nothing
itemUrl = AliPayUrl
'Response.Redirect AliPayUrl   '跳转到支付页
							
							%>
          <div style="line-height:5px;">&nbsp;</div>
          <table width="98%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td align="left" valign="top" style="line-height:180%; font-size:15px; padding-left:12px"> 支付宝帐号：<br />
                <span class="fntF00">mtc616@yeah.net</span></td>
              <td width="200" align="center"><a href="<%=itemUrl%>" target="_self"><img src="alipay/alipay_bwrx.gif" width="180" height="60" border="0" /></a></td>
            </tr>
          </table>
          </td>
      </tr>

    </table>
      </div></td>
  </tr>
</table>
<script language="javascript">
function caFee(e,xID)
{
  var fm = eval("document.fm01"+xID); // PAdult,PChild,PBaby,nFee,
  var ev = e.value; //M1_<%=SubID%>
  if ( isNaN(ev) || ev.indexOf(".")>=0 ){
    e.value = 0;
	alert("请填写数字！");
  }
  if (fm.nAdult.value==0){
    fm.nAdult.value = 1
	alert("成年数 至少为1！");
  }
  if((fm.nAdult.value*1+fm.nChild.value*1)>fm.NTotal.value*1){
    e.value = 0;
	alert("成人+小孩,最多可报"+fm.NTotal.value+"人!");
  }
  
  fm.nFee.value = fm.PAdult.value*fm.nAdult.value + fm.PChild.value*fm.nChild.value + fm.PBaby.value*fm.nBaby.value;
  M9_All.innerHTML = fm.nFee.value;
  eval("M1_"+xID).innerHTML = fm.nFee.value;
}
</script>
</body>
</html>
