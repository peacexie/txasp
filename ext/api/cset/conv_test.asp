<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test</title>
</head>
<body>
<!--#include file="conv_func.asp"-->

<%  

s1 = "aa亿美软通，张工，123，abc"
response.write "<br>"&cUrl_gb2u8(s1)  
response.write "<br>"
response.write "<br>"&Server.URLEncode(s1)  

'FF,Google 656(2:1312; 6:)
'IE8:454(2:908; 6:2724)
'IE7:Timout
'IE6:222(2:444; 6:) 
'260(4),195(3x65)

s2 = "测试" : s3 = ""
for ixx=1 To 328 '210'256'250'
  s3 = s3&s2
next
Response.Write "<hr>"&Len(s3)&s3
s3 = Server.URLEncode(s3)
s4 = Request("s4")
Response.Write "<hr>"&Len(s4)&s4

'Max:454 454*6 = 2724
'http://localhost:240/u/demo/ext/cset/GB2312-UTF8.asp?s4=
'12345678901234567890123456789012345678901234567890123456 (56)
%>
<br />

xx
<a href='?s4=<%=s3%>'>Send Test</a>
xx


<%

Function c_x1(xStr)
  Session.Codepage = 950
  t = Server.URLEncode(xStr)
  'Response.Write xStr
  Session.Codepage = 65001'还原当前会话编码
  c_x1 = t
End Function

s00 = "Test--代刚--?#&%abc@_" 
s00 = c_x1(s00) :Response.Write s00
s00 = cUrl_Decode(s00)
Response.Write s00
For ij=1 To Len(s00)
  ch = Mid(s00,ij,1)
  c5 = c_x1(ch)
  'c5 = Replace(c5,"%","")
  'cu = cNum_U8toU(c5)
  'cb = cBin_s2b(ch, "big5")
  'cb2 = cBin_b2s(midB(cb,1),"big5") 
  Response.Write "<br>"&ch&" - "&c5
  'response.binaryWrite cb
Next

Response.Write "<br><br>"
h = "6D4B"
a = cNum_U8toU(h)
c = ChrW(a)
Response.Write "<br> - "&a&" - "&c
Response.Write "<br><br>"

h = "44"
a = cNum_U8toU(h)
c = ChrW(a)
Response.Write "<br> - "&a&" - "&c

h = "424"
a = cNum_U8toU(h)
c = ChrW(a)
Response.Write "<br> - "&a&" - "&c

h = "E8BDAC"
a = cNum_U8toU(h)
c = ChrW(a)
Response.Write "<br> - "&a&" - "&c

ch = "D" 
cAsc = AscW(ch) :If cAsc<0 Then cAsc=cAsc+65536
cUrl = Server.URLEncode(ch)
cre = ChrW(cAsc)
cu8 = cNum_utou8(cAsc)
cu8b = c16to2(cu8)
Response.Write "<br>1."&cre&" - "&cAsc&" - "&cUrl
Response.Write "<br>2."
Response.Write Hex(cAsc)
Response.Write " - "&cu8b

ch = "Ф" 
cAsc = AscW(ch) :If cAsc<0 Then cAsc=cAsc+65536
cUrl = Server.URLEncode(ch)
cre = ChrW(cAsc)
cu8 = cNum_utou8(cAsc)
cu8b = c16to2(cu8)
Response.Write "<br>1."&cre&" - "&cAsc&" - "&cUrl
Response.Write "<br>2."
Response.Write Hex(cAsc)
Response.Write " - "&cu8b

ch = "转" 
cAsc = AscW(ch) :If cAsc<0 Then cAsc=cAsc+65536
cUrl = Server.URLEncode(ch)
cre = ChrW(cAsc)
cu8 = cNum_utou8(cAsc)
cu8b = c16to2(cu8)
Response.Write "<br>1."&cre&" - "&cAsc&" - "&cUrl
Response.Write "<br>2."
Response.Write Hex(cAsc)
Response.Write " - "&cu8b


'UNICODE  - UTF-8 
'00000000 - 0000007F 0xxxxxxx 
'00000080 - 000007FF 110xxxxx 10xxxxxx 
'00000800 - 0000FFFF 1110xxxx 10xxxxxx 10xxxxxx 
'               8F6C 11101000 10111101 10101100
'00010000 - 001FFFFF 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx 
'00200000 - 03FFFFFF 111110xx 10xxxxxx 10xxxxxx 10xxxxxx 10xxxxxx 
'04000000 - 7FFFFFFF 1111110x 10xxxxxx 10xxxxxx 10xxxxxx 10xxxxxx 10xxxxxx 


Response.Write "<br>"
Response.Write "<br>"
'ascw(永)  : 6C38
'url(utf-8): %E6%B0%B8

dim str:str="adodb.stream 实现 二进制与字符串的互转 By Shawl.qiu" 
response.write "字符串转二进制 response.binaryWrite cBin_s2b(str, ""utf-8""):<br/>" 
response.binaryWrite cBin_s2b(str, "utf-8") 
'response.Write cBin_s2b(str, "utf-8") 
response.write "<p/>二进制转字符串 response.write cBin_b2s(midB(cBin_s2b(str, ""utf-8""),1),""utf-8"")<br/>" 
response.write cBin_b2s(midB(cBin_s2b(str, "utf-8"),1),"utf-8") 


response.write "<hr/>"


'用途:將UTF-8編碼漢字轉換為GB2312碼，兼容英文和數字 
'版權:雖說是原創，其實也參考了別人的部分算法 
'用法:Response.write cUrl_ut2gb("%E9%83%BD%E5%B8%82%E6%83%85%E7%B7%A3 %E6%98%9F%E5%BA%A7") 
'["岁月联盟"提供]
'Response.Write "-版權:雖說是原創"
s1 = "Test%E6%B5%8B%E8%AF%95%E5%AD%97%E7%AC%A6%E4%B8%B2%E7%9A%84%E4%BA%92%E8%BD%AC"
Response.Write s1
'Response.Write cUrl_Chinese(s1)
Response.Write cUrl_ut2gb(s1)
Response.Write cUrl_Chinese("%E6%B5%8B")
Response.Write cUrl_ut2gb("%E6%B5%8B")


response.write "<hr/>"


Response.Write "<pre>"
org = "Test--测试--?#&%abc@_"
enc = cUrl_Escape(org)
un = cUrl_UnEsc(enc) 
url = Server.URLEncode(org)
q = Request("q")
Response.Write vbcrlf&"<br>org:"&org&"-----<a href='?'>"&org&"</a>"
Response.Write vbcrlf&"<br>enc:"&xxx&"-----<a href='?q="&enc&"'>"&enc&"</a>"
Response.Write vbcrlf&"<br>un-:"&un
Response.Write vbcrlf&"<br>url:"&xxx&"-----<a href='?q="&url&"'>"&url&"</a>"
Response.Write vbcrlf&"<br>q--:"&q

Response.Write vbcrlf&"<br>"
Response.Write vbcrlf&"<br>测试-測試-望眼:"&cUrl_Escape("测试-測試-望眼")
Response.Write vbcrlf&"<br><a href='enc_b5.asp?q="&enc&"'>b5_enc</a> | <a href='enc_gb.asp?q="&url&"'>gb_url</a>"
Response.Write vbcrlf&"<br><a href='enc_gb.asp?q="&enc&"'>gb_enc</a> | <a href='enc_b5.asp?q="&url&"'>b5_url</a>"
Response.Write vbcrlf&"<br><a href='?q="&q&"'>q_org</a> | <a href='?q="&cUrl_Escape(q)&"'>q_enc</a> "
Response.Write vbcrlf&"<br><a href='enc_b5.asp?q="&cUrl_Escape(q)&"'>b5_enc</a> | <a href='enc_gb.asp?q="&cUrl_Escape(q)&"'>gb_enc</a>"
Response.Write vbcrlf&"<br><a href='enc_b5.asp?q="&Server.URLEncode(q)&"'>b5_url</a> | <a href='enc_gb.asp?q="&Server.URLEncode(q)&"'>gb_url</a>"

Response.Write vbcrlf&"</pre>"

%> 


</body>
</html>
