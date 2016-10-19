<%
	'版本：2.0
	'日期：2009-01-05
	'作者：支付宝公司销售部技术支持团队
	'联系：0571-26888888
	'版权：支付宝公司
%>
<!--#include file="alipayto/alipay_payto.asp"-->
<%

if request("action")="sendok" then 
dim service,partner,sign_type,out_trade_no
dim t1,t4,t5,key
dim AlipayObj,itemUrl
	t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'支付接口

	t4				=	"images/alipay_bwrx.gif"		'支付宝按钮图片
	t5				=	"推荐使用支付宝付款"						'按钮悬停说明
	
	service         =   "single_trade_query"
	partner			=	"2088101707183363"		'partner合作伙伴ID（签约支付宝账号，商家服务可以查询到合作id和安全校验码）
	sign_type       =   "MD5"
   	out_trade_no    =    request("gross")  '要查询的外部订单号码
	key             =    "rcaj004vn57t6ndfb6s7r7b37y62oa7j" '签约账户对应的安全校验码，


	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(t1,t4,t5,service,partner,sign_type,out_trade_no,key)
    'response.redirect itemUrl
	url=itemUrl
	Dim http,xml
	Set http=Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET",url,False
	http.send
	Set xml=Server.CreateObject("Microsoft.XMLDOM")
	xml.Async=true
	xml.ValidateOnParse=False
	xml.Load(http.ResponseXML)
	
	
	
	Set UserData=xml.getElementsByTagName("trade")  ' 节点的名称
	
	if isnull(xml.getElementsByTagName("alipay") ) then
		response.Write("读取失败")
		response.End()
  	else
	
    	for j=0 to UserData.item(i).childnodes.length-1
		
    		Response.Write UserData.item(0).childnodes(j).text &"<br>*************"
			'mystr = SPLIT(mystr, "^")
   		next
	end if
	response.Write("解析成功")

Set http=Nothing
Set xml=Nothing

%>
</td>

<%
else
%>
<form action="Search.asp?action=sendok" method="post" name="myform">
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=450>
<tr><td align=center width=50% height=120><p>外部订单号码：
    <input name="gross" type="text" value="" />
    <input name="addpost" type="hidden" value="addpost" />
  </p>
  <p>
      <input name="submit" type="submit" value="提交" />
    </p></td>
</tr>

<tr><td colspan=2 align=center>本站使用支付宝支付平台，在线实时支付<br><br>
<a href='http://www.alipay.com'  target="_blank"><img src='https://img.alipay.com/pimg/logo.gif' border=0></a>
</td></tr>
</table>	
<%End if%>