<%
a = Request("a")
u = Request("u")
p = Request("p")
Response.Write "a=["&a&"],u=["&u&"],p=["&p&"],"
%>


<pre>
<complexType name="getBalanceResponse">
  <complexContent>
    <restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
      <sequence>
        <element name="return" type="{http://www.w3.org/2001/XMLSchema}double"/>
      </sequence>
    </restriction>
  </complexContent>
</complexType>

<complexType name="getBalance">
  <complexContent>
    <restriction base="{http://www.w3.org/2001/XMLSchema}anyType">
      <sequence>
        <element name="arg0" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
        <element name="arg1" type="{http://www.w3.org/2001/XMLSchema}string" minOccurs="0"/>
      </sequence>
    </restriction>
  </complexContent>
</complexType>

�߷����,ASP��ε���web services?(Ŀǰֻ֪��WSDL)

http://help.cj.com/en/web_services/web_services.htm
����web services��˵��.

https://product.api.cj.com/wsdl/version2/productSearchServiceV2.wsdl
����WSDL�ļ�.

������aspʹ��xmlhttp����web services��ô����?

�������ϲ鵽����������������Ĵ���:


    ' Client invoke WebService use HTTP POST request and response
    Function vbGetHelloWorld_HTTPPOST(i)
        URL = "http://localhost/WebServicesTest/Default.asmx/HelloWorld"
        Params = "i=" & i ' Set postback parameters
        
        Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
        xmlhttp.Open "POST",URL,False
        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp.setRequestHeader "Content-Length",LEN(Params)
        xmlhttp.Send(Params)
        Set x =   xmlhttp.responseXML
        
        if xmlhttp.Status = 200 then
            response.Write(x.childNodes(1).text)
            'alert(xmlhttp.Status)
            'alert(xmlhttp.StatusText)
        end if
        Set xmlhttp = Nothing
    End Function




