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
AliPay.Service     = "trade_create_by_buyer"  'create_direct_pay_by_user  (即时到账) ，  trade_create_by_buyer 标准双接口
'AliPay.Charset       = "UTF-8"    '编码 默认UTF-8
'**************以上部分建议在类中设置**************
AliPay.Subject       = "测试商品"    '商品名称
AliPay.Body          = "商品描述"    '商品描述
AliPay.Price         = "0.01"        'price商品单价  0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
AliPay.Quantity      = "1"           '商品数量
'**************以下部分可以在类中设置**************
AliPay.Out_Trade_No  = Year(Now())&Month(Now())&Day(Now())&Hour(Now())&Minute(Now())&Second(Now())'按时间获取的订单号
AliPay.Discount      = "0"             '商品折扣
AliPay.Paymethod     = "directPay"     '赋值:bankPay(网银);cartoon(卡通); directPay(余额)
'///////////////////////以下是标准双接口设置的//////////////////////////////
'标准双接口无论如何第一组都要保留，不可以为空
'第一组
AliPay.logistics_fee	    =	"0.00"			'物流配送费用  0.00
AliPay.logistics_payment  =	"BUYER_PAY"	   '物流配送费用付款方式：SELLER_PAY(卖家支付)  BUYER_PAY(买家支付)  BUYER_PAY_AFTER_RECEIVE(货到付款)
AliPay.logistics_type	    =	"EXPRESS"		'物流配送方式：POST(平邮)、EMS(EMS)  EXPRESS(其他快递)
'如果需要多添加几组物流方式，可以增加第二组物流参数,如果不需要，可以为空
AliPay.logistics_fee_1    =	""			'物流配送费用  0.00
AliPay.logistics_payment_1=	""	'物流配送费用付款方式：SELLER_PAY(卖家支付)  BUYER_PAY(买家支付)  BUYER_PAY_AFTER_RECEIVE(货到付款)
AliPay.logistics_type_1	=	""		'物流配送方式：POST(平邮)  EMS(EMS)  EXPRESS(其他快递)
'如果需要多添加几组物流方式，可以增加第三组物流参数,如果不需要，可以为空
AliPay.logistics_fee_2    =	""			'物流配送费用  0.00
AliPay.logistics_payment_2=	""	'物流配送费用付款方式：SELLER_PAY(卖家支付)  BUYER_PAY(买家支付)  BUYER_PAY_AFTER_RECEIVE(货到付款)
AliPay.logistics_type_2	=	""		'物流配送方式：POST(平邮)  EMS(EMS)  EXPRESS(其他快递)
'以下是可选的参数 如果没有可以为空。注意：姓名、联系地址和邮政编码 这三项要么都为空，要么都不能为空。
AliPay.ShowUrl       =	""    '商品的演示地址
AliPay.receive_name    =	""   '收货人姓名
AliPay.receive_address =	""   '收货人地址
AliPay.receive_zip     =	""   '邮编5 位戒6 位数字组成
AliPay.receive_phone   =	""   '收货人电话
AliPay.receive_mobile  =	""   '收货人手机 必须是11 位数字
AliPay.buyer_email   =	""   '买家的支付宝账号
'///////////////////////以上是标准双接口设置的//////////////////////////////
AliPayUrl = AliPay.GetUrl()
Set AliPay = Nothing
Response.Redirect AliPayUrl   '跳转到支付页
%>