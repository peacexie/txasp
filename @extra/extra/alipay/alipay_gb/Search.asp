<%
	'�汾��2.0
	'���ڣ�2009-01-05
	'���ߣ�֧������˾���۲�����֧���Ŷ�
	'��ϵ��0571-26888888
	'��Ȩ��֧������˾
%>
<!--#include file="alipayto/alipay_payto.asp"-->
<%

if request("action")="sendok" then 
dim service,partner,sign_type,out_trade_no
dim t1,t4,t5,key
dim AlipayObj,itemUrl
	t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'֧���ӿ�

	t4				=	"images/alipay_bwrx.gif"		'֧������ťͼƬ
	t5				=	"�Ƽ�ʹ��֧��������"						'��ť��ͣ˵��
	
	service         =   "single_trade_query"
	partner			=	"2088101707183363"		'partner�������ID��ǩԼ֧�����˺ţ��̼ҷ�����Բ�ѯ������id�Ͱ�ȫУ���룩
	sign_type       =   "MD5"
   	out_trade_no    =    request("gross")  'Ҫ��ѯ���ⲿ��������
	key             =    "rcaj004vn57t6ndfb6s7r7b37y62oa7j" 'ǩԼ�˻���Ӧ�İ�ȫУ���룬


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
	
	
	
	Set UserData=xml.getElementsByTagName("trade")  ' �ڵ������
	
	if isnull(xml.getElementsByTagName("alipay") ) then
		response.Write("��ȡʧ��")
		response.End()
  	else
	
    	for j=0 to UserData.item(i).childnodes.length-1
		
    		Response.Write UserData.item(0).childnodes(j).text &"<br>*************"
			'mystr = SPLIT(mystr, "^")
   		next
	end if
	response.Write("�����ɹ�")

Set http=Nothing
Set xml=Nothing

%>
</td>

<%
else
%>
<form action="Search.asp?action=sendok" method="post" name="myform">
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=450>
<tr><td align=center width=50% height=120><p>�ⲿ�������룺
    <input name="gross" type="text" value="" />
    <input name="addpost" type="hidden" value="addpost" />
  </p>
  <p>
      <input name="submit" type="submit" value="�ύ" />
    </p></td>
</tr>

<tr><td colspan=2 align=center>��վʹ��֧����֧��ƽ̨������ʵʱ֧��<br><br>
<a href='http://www.alipay.com'  target="_blank"><img src='https://img.alipay.com/pimg/logo.gif' border=0></a>
</td></tr>
</table>	
<%End if%>