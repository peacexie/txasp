<%@ LANGUAGE = VBScript CodePage = 65001%>
<%
Option Explicit
Response.Buffer = True
Session.CodePage = 65001
'Session.LCID = 2057
Response.charset = "utf-8"
response.expires=-1
response.expiresabsolute=now()-1
response.cachecontrol="no-cache"
%>
<!--#include file="AliPay_Class.asp"-->
<%
Dim AliPay, AliPayUrl
Set AliPay = New Qlwz_AliPay
'**************以下部分建议在类中设置**************
'AliPay.PartnerID     = ""    '填写对应支付宝账户的合作者身份ID
'AliPay.SellerEmail   = ""    '请填写支付宝签约帐户
'AliPay.SignCode      = ""    '填写对应支付宝帐户的安全校验码
'AliPay.NotifyUrl     = "http://......./Notify_Url.asp"    '交易过程中服务器通知的页面,,,绝对路径，要求可以访问
'AliPay.ReturnUrl     = "http://......./Return_url.asp"    '付完款后跳转的页面,,,绝对路径，要求可以访问
'AliPay.Service     = "create_direct_pay_by_user"  'create_direct_pay_by_user  (即时到账) ，  trade_create_by_buyer 标准双接口
'AliPay.Charset       = "UTF-8"    '编码 默认UTF-8
'**************以上部分建议在类中设置**************
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