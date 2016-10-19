<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Demo 0</title>
</head>
<body>
<a href="AliPay_Demo1.asp" >即时到账create_direct_pay_by_user.asp</a><br>
<a href="AliPay_Demo2.asp" >标准双接口trade_create_by_buyer.asp</a><br>
<a href="AliPay_Demo3.asp" >即时到账改进</a><br>
<br>
<a href="AliPay_Notify.asp?id=MTC200912290013&ActTest=(PeaceTest)" >Test Peace</a><br>
<a href="AliPay_Return.asp" >AliPay_Return.asp</a>
<hr />
<a href="?a=1x&b=2x&c=3x&d=4x">4aTest</a>,
<a href="?b=1t&a=2t&c=3t&d=4t">4bTest</a>,
<a href="?d=1c&c=2c&b=3c&a=4c">4cTest</a>,
<a href="?c=1w&d=2w&b=3w&a=4w">4dTest</a>,
<hr />
<a href="?a=1x&b=2x">2aTest</a>,
<a href="?b=1t&a=2t">2bTest</a>,
<a href="?a=1x">1aTest</a>,
<a href="?">0aTest</a>,
<hr />
<a href="?a=1x&b=2x&c=3x&d=4x&e=5hh">5aTest</a>,
<a href="?b=1t&a=2t&c=3t&d=4t&e=5hh&f=6xxx">6bTest</a>,
<a href="?d=1c&c=2c&b=3c&a=4c&e=5hh&f=6xxx&g=Peace">7cTest</a>,
<a href="?c=1w&d=2w&b=3w&a=4w&e=5hh&f=6xxx&h=Shirley&i=i">8dTest</a>,
<hr /> 
<a href="?c=1w制5551造&d=2w3储如1733高效管71133869理&b=3w高效管71133869理&a=4w高效管71133869理&e=5hh管71133869理&f=6管71133869理xxx&h=Shirley&i=i">8dTest</a>,
<a href="?c=1w仓72963&d=2w3储如1733高效管71133869理&b=3w高效管71133869理&a=4w高效管71133869理&e=5hh管71133869理&f=6xx管71133869理x&h=Sh管71133869理irley&i=i管71133869理管71133869理">8dTest</a>,
<a href="?c=1w仓72963&d=2w3储如1733高效管71133869理&b=3w高效管71133869理&a=4w高效管71133869理&e=5hh管71133869理&f=6xx管71133869理x&h=Shi管71133869理rley&i=i管71133869理">8dTest</a>,
<a href="?c=1w仓72963&d=2w3储如1733高效管71133869理&b=3w高效管71133869理&i=4w高效管71133869理&e=5hh管71133869理&f=6xx管71133869理x&h=Shir管71133869理ley&a=管71133869理i">8dTest</a>,
<hr />
<%

'/////////////////////////////////
'/////////////////////////////////
Dim tm1,tm2,st1,st2,st3,st4,n
tm1 = Timer()
For n = 1 To 1 '300

'st1 = rForm()
st1 = rQuery()
st2 = rSort(st1,"")
st3 = rSort(st1,"Items")
st4 = rSort(st1,"Vals")

Next
tm2 = Timer()
Response.Write "<br>All Time: "&tm2-tm1


Response.Write "<br>t1:  ("&st1&")"
Response.Write "<br>t2:  ("&st2&")"
Response.Write "<br>t3:  ("&st3&")"
Response.Write "<br>t4:  ("&st4&")"
'Response.Write "<pre>"
'Response.Write "</pre>"

'/////////////////////////////////
'/////////////////////////////////


Function rSort(xStr,xType) 'Type:Items,Vals,Both
Dim s0,s1,s2,a0,a1,a2
  s0 = xStr&"" 
  If inStr(s0,vbcrlf)>0 Then
	a0 = Split(s0,vbcrlf)
	a1 = Split(a0(0),"^")
	a2 = Split(a0(1),"^")
	s0 = rSArr(a1,a2)
  Else
    s0 = "^" 'xType="Both"
  End If
  If xType="Items" Then
    a3 = Split(s0,"^")
	s0 = a3(0)
  ElseIf xType="Vals" Then
    a3 = Split(s0,"^")
	s0 = a3(1)
  End If
rSort = s0
End Function
Function rSArr(xA1,xA2)
Dim i,j,a1,b1,a2,b2,s1,s2 : s1="" : s2=""
  For i = uBound(xA1) To 0 Step -1 '0 TO uBound(xA1)
  For j = 0 TO i-1
    a1 = xA1(j) : a2 = xA1(j+1)  
	b1 = xA2(j) : b2 = xA2(j+1)
	If a1>a2 Then
	  xA1(j) = a2  : xA1(j+1) = a1 
	  xA2(j) = b2  : xA2(j+1) = b1
	End If
  Next
  Next
  For i = 0 TO uBound(xA1)
    s1 = s1 & xA1(i)
	s2 = s2 & xA2(i)
  Next
rSArr = s1&"^"&s2
End Function

Function rForm()
Dim iItem,s1,s2 : s1="" : s2=""
  For Each iItem in Request.Form
    s1 = s1&"^"&iItem
    s2 = s2&"^"&rFill(Request.Form(iItem))
  Next 
rForm = s1&vbcrlf&s2
End Function
Function rQuery()
Dim iItem,s1,s2 : s1="" : s2=""
  For Each iItem in Request.QueryString
    s1 = s1&"^"&iItem
    s2 = s2&"^"&rFill(Request.QueryString(iItem))
  Next 
rQuery = s1&vbcrlf&s2
End Function
Function rFill(xStr)
Dim s : s = xStr&""
  s = Replace(s,"'","")
  s = Replace(s,"^","")
  s = Replace(s,vbcrlf,"")
  s = Replace(s,vbcr,"")
  s = Replace(s,vblf,"")
rFill = s
End Function


%>
</body>
</html>
