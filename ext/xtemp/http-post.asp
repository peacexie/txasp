<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<!--#include file="../../sadm/func2/func_obj.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test Code Demo</title>
</head>
<body>

<%

Function retuText(send_URL,user,pw) 
  PostStr = "?a=b&u="&user&"&p="&pw 
  Response.Write "<br>"&Len(PostStr)&"<br>"
  Set Retrieval = Server.CreateObject("Msxml2.xmlHttp") 'Microsoft.XMLHTTP
  With Retrieval 
	.Open "POST",send_URL,False," ", " " 
	.setRequestHeader "Content-Length",len(PostStr) 
	.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded" 
	.Send(PostStr)
	retuText = .ResponseText 
  End With 
  Set Retrieval = Nothing 
End Function 

s = ""
for i=1 to 32012 '32K
  Randomize
  p = 65+Int(26*Rnd()) ' 0 ~ (rMax-1)
  s = s&Asc(p)&""
next
Response.Write "<br>Len(s):"&len(s)

a1 = "http://www.dg.gd.cn/dgnews/view.asp?ID=132DBFA55DB7BATE133KPWKH"
a2 = "http://localhost:240/u/demo/ext/xtemp/http-get.asp"
url = a2 
user = "test12" 
pw = "pw123423"&s 

te=retuText(url,user,pw) 
Response.Write "<br>Len(te):"&Len(te)&"<br>"
Response.Write "<br>"&vbcrlf&"<!--"&vbcrlf&te&vbcrlf&"--><br>"
Response.Write "<br>"&Now()
Response.End()

'JmailServer = "smtp.gmail.com"
'JmailID = "elifebike@gmail.com"
'JmailPW = "19800203"
%>

<%

echo(Request.Servervariables("HTTP_USER_AGENT"))

'yyyy-mm-dd hh:nn:ss
'  99 12 31 23 59 59
'yyyymmdd hhnnss

a = Hex(0101)   :echo(a)
a = Hex(121231) :echo(a)
a = Hex(991231) :echo(a)
a = Hex(235959) :echo(a)
a = Hex(1001231) :echo(a)
echo("")

a = Hex(000101) :echo(a)
a = Hex(991231) :echo(a)
a = Hex(19001231) :echo(a)
a = Hex(20001231) :echo(a)
a = Hex(99991231) :echo(a)
echo("")
a = Hex(000000) :echo(a)
a = Hex(235959) :echo(a)

echo("")
d1 = Get_yyyymmdd(xDate)
d2 = Get_hhmmss()
a = Hex(d1) :echo(a)
a = Hex(d2) :echo(a)

echo("")

echo(Date())
echo(Time())
echo(Day(Date()))
echo(Hex("1231235959")) '2100
'echo(Hex("1231235959")) '2100
echo("")
xTime = Now()
sm = Mid(Pub_IDStr,DatePart("m",xTime)+1,1)
sd = Mid(Pub_IDStr,DatePart("d",xTime)+1,1)
'sh = Mid(Pub_IDStr,DatePart("h",xTime),1)
'mm = Mid(Pub_IDStr,DatePart("n",xTime),1)
sh = DatePart("h",xTime)
mm = DatePart("n",xTime)
echo(sm&sd&sh&mm)
echo(Hex(1100))
echo(Hex(1232359))

'Response.End() DatePart("n",Now())
'Call Send_jMail("Name","xpigeon@163.com","Test "&Now,"xBody "&Now,"gb2312") 

s = ""
For i=0 To 65535
  's = s&ChrW(i)
Next
'Call File_Add2("./ch.txt",s,"utf-8")

s = "ss"
a = Split(s,"")
Response.Write uBound(a)
Response.Write "<br>"

c = Int(2/3)
Response.Write "<br>"&c
c = Int(7/3)
Response.Write "<br>"&c


c = Int(12/3)
Response.Write "<br>"&c
c = Int(12/5)
Response.Write "<br>"&c
c = Int(12/11)
Response.Write "<br>"&c


fiName = "aaaaaaaaa.Peace!Bak"
Response.Write "<br>"&Mid(fiName,InStrRev(fiName,"."),12)


xUrl = "/u/demo/upfile/dtinf/2011/1A/8WDS.88XVU/1A8WFU_9YMRT~copy.jpg"
Response.Write "<br>"&Mid(xUrl,InStrRev(xUrl,"/")+1) 

Response.Write "<br>"&OutSFlag("13452","1","2")
Response.Write "<br>"&OutSFlag("1134522","11","22")


tTimer1 = Timer()


set rs = Server.CreateObject("ADODB.recordset")
rs.Open "SELECT SysID, SysName FROM AdmSyst", conn
  str = rs.GetString(,,"</td><td>","</td></tr>"&vbcrlf&"<tr><td>","&nbsp;")
rs.close
set rs = Nothing
str = Replace(str&"^",""&vbcrlf&"<tr><td>^","")
%>
<table border="1">
<tr><td><%Response.Write(str)%>
</table>
<%

Response.End()


Set rs = Server.CreateObject("ADODB.recordset") 'New ADODB.Recordset
rs.CursorLocation = adUseClient

rs.Open "Select * From AdmSyst", conn, 1, 1
rs.PageSize = 5
rs.AbsolutePage = 3
lngPages = rs.PageCount
lngCurrentPage = 1
Response.Write "<br> cur:"&rs.CursorLocation&" <br>cnt:"&rs.PageCount


Response.Write Request("AA")
%>

<form action="" method="get">
<textarea name="AA" cols="46" rows="5" id="AA">< OKOKOK >
&lt;223322&gt;
</textarea>

<input type="submit" name="button" id="button" value="Submit" />
</form>
<pre>

<%




sys27_Rand = Rnd_Base("5678",9)&Rnd_Base("",64)
Dim sys27_Rnd(10)
For i = 1 To 9
 sys27_Rnd(i)=Mid(sys27_Rand,i*6,Mid(sys27_Rand,i,1))
Next


Response.Write vbcrlf&"<br>-:1234567890123456789012345678901234567890"
Response.Write vbcrlf&"<br>-:"&sys27_Rand
For i = 1 To 9 
 Response.Write vbcrlf&"<br>"&i&":"&sys27_Rnd(i)
Next

%>
</pre>
</body>
</html>
