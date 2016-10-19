<%@ LANGUAGE = VBScript CodePage = 65001%>

<!--#include file="AliPay_Class.asp"-->
<%
Dim AliPay, AliPayUrl
Set AliPay = New Qlwz_AliPay

AliPay.Subject       = "1111商品"    '商品名称
AliPay.Body          = "1111商品描述"    '商品描述
AliPay.Price         = "0.01"        'price商品单价  0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
AliPay.Quantity      = "1"           '商品数量
'**************以下部分可以在类中设置**************
AliPay.Out_Trade_No  = Year(Now())&Month(Now())&Day(Now())&Hour(Now())&Minute(Now())&Second(Now())'按时间获取的订单号
AliPay.Discount      = "0"             '商品折扣
AliPay.Paymethod     = "directPay"     '赋值:bankPay(网银);cartoon(卡通); directPay(余额)

AliPayUrl = AliPay.GetUrl()
Set AliPay = Nothing
Response.Redirect AliPayUrl   '跳转到支付页
%>