<!--#include file="alipayto/alipay_payto.asp"-->
<%
   shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	
	subject			=	"���Զ�����"		'��Ʒ����
	body			=	"a"		'body			��Ʒ����
	out_trade_no    =   dingdan       
	price		    =	"0.01"				'price��Ʒ����			0.01��50000.00
    quantity        =   "1"               '��Ʒ����,����߹��ﳵĬ��Ϊ1
	discount        =   "0"               '��Ʒ�ۿ�
    seller_email    =    seller_email   '���ҵ�֧�����ʺ�
	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=450>
<tr><td align=center width=50% height=120>
<a href="<%=itemUrl%>" target="_blank"><img src="alipay_bwrx.gif" border="0"></a>
</td>

</tr>
<tr><td colspan=2 align=center>��վʹ��֧����֧��ƽ̨������ʵʱ֧��<br><br>
<a href='http://www.alipay.com'  target="_blank"><img src='https://img.alipay.com/pimg/logo.gif' border=0></a>
</td></tr>
</table>	