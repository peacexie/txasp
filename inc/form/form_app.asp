<%

'// xOut为超时分钟;
'// true = 超时退出
Function A99ChkTime(xTime,xOut) 
  'nNow = Now() 'strtotime(date("Y-m-d H:i:s")) '//'2005-10-12 20:02:40'
  'nTim = ceil($nNow-strtotime($xTime)) '//echo $nTim; 
  'echo xTime
  nTime = DateDiff("s",xTime,Now())
  If nTime>xOut*60 Then
	A99ChkTime = true
  Else
	A99ChkTime = false 
  End If

End Function

'xOut为超时分钟; xSess="" 不认证Session
'app30Time = Request(App30Code&"0")
'app30TSN = Request(App30Code&"1") 
'app30TStr = Request(App30Code&"2")
'app30Tab = Split(app30TStr,"-")
'If Get_A30Chk(app30Time,app30TSN,60)<>"" Then
Function Get_A30Chk(xTim,xTSN,xOut,xSess)
  Dim nTim,r :r="" 'xOut:1H=60,1D=1440,
  nTim = Abs(DateDiff("n",xTim,Now())) ':echo(nTim)
  If nTim>xOut Then
    r = "TOut" 'Tim超时错误
  ElseIf xTSN<>MD5_32(App30Code&"@"&xTim) Then
    r = "TEnc" 'TSN认证错误
  End If 
  If xSess<>"" And Session(App30Code)&""<>xTSN Then 
    r = "TSes" '认证Session
  End If
  'echo r
  Get_A30Chk = r
End Function

'//xSess="" 不认证Session
'//$app99Tim = date("Y-m-d H:i:s");
'//$app30Arr = App30Set($app99Tim,"s","");
'//$app30Tab = explode("-",$app30Arr[2]);
Function App30Set(xTime,xSess,xType)
  'global $App30Code; 
  Dim sTim,sTSN,sEID,sTab,x,y,z,a,b,c,i,s,p
  If xTime="" Then
    sTim = Now()
  Else
    sTim = xTime
  End If
  sTSN = MD5_32(App30Code&"@"&sTim) 
  sEID = Get_IDEnc(sTSN,0)&sTSN
  sTab = ""
  x = DatePart("h",sTim) :a = (x mod 12)+97 '//Asc(97=a)
  y = DatePart("n",sTim) :b = (x mod 3)+3
  z = DatePart("s",sTim) :c = (x mod 8)+8
  s = uCase(Mid(sTSN,18,c)) ' strtoupper(substr($sTSN,18,$c));
  For i=1 To 12
    p = i*3-1
	sTab = sTab&Chr(a+i)&lCase(Mid(sEID,p,b))&"_"&s&"-"
  Next
  if xSess<>"" Then Session(App30Code)=sTSN
  App30Set = Split(sTim&"@"&sTSN&"@"&sTab,"@")
End Function

'xSess="" 不认证Session
'app30Arr = Get_A30SN("","s","")
'app30Tab = Split(app30Arr(2),"-")
Function Get_A30SN(xTime,xSess,xType) 
  Dim ra(2),sTim,sTSN,sTab,x,y,z,s,p,a,b,c,i
  sTab = "" :sTim = Now() 
  If xTime<>"" Then sTim = xTime    
  sTSN = MD5_32(App30Code&"@"&sTim) 
  sEID = Get_IDEnc(sTSN,0)&sTSN
  x = DatePart("h",sTim) :a = (x Mod 12)+97 'Asc(97=a)
  y = DatePart("n",sTim) :b = (y Mod 3)+3 '3~5
  z = DatePart("s",sTim) :c = (z Mod 6)+8 '8~13
  s = uCase(Mid(sTSN,18,c)) 'echo(Chr(a)&b&c)
  For i=1 To 13 
    p = i*3-1
	sTab = sTab&Chr(a+i)&lCase(Mid(sEID,p,b))&"_"&s&"-"
  Next
  'sTab = Left(sTab,Len(sTab)-1)
  ra(0)=sTim :ra(1)=sTSN :ra(2)=sTab
  If xSess<>"" Then Session(App30Code)=sTSN
  Get_A30SN = ra
End Function

'n为超时分钟; 
Function Get_AChk(xStr,xMin)
  Dim i,s
  For i=0 To xMin
    s = Get_AMin(-1*i)
	If s=xStr Then
	  Get_AChk = true
	  Exit Function
	End If
  Next
  Get_AChk = false
End Function 
'n为时间偏移; 
Function Get_AMin(xMin)
  Dim sid,tim,stn,stc
  sid = Session.SessionID
  tim = DateAdd("n",xMin,Now())
  stn = DatePart("m",tim)
  stn = stn&DatePart("d",tim)
  stn = stn&DatePart("h",tim)
  stn = stn&DatePart("n",tim)
  'stn = stn&DatePart("s",tim)
  stc = Int(Right(sid,8))+Int(Right(stn,8))
  stc = Hex(stc)
  Get_AMin = stc
End Function 
'// 检查是否为每小时Get_AHour的加密串,n为超时分钟; 
'// Demo IF Not Get_AHChk(Request(App28Code),30) Then 
Function Get_AHChk(xStr,xMin)
  Dim mm,s0,m0,s1,s2
  mm = Int(Mid(Get_hhmmss(),3,2)) '分钟; eg.34
  s0 = Left(xStr,32)
  m0 = Int(Right(xStr,2)) '分钟; eg.12
  If mm<m0 Then 
    mm=mm+60
    s1 = Get_AHour(60) '上1小时的
  Else
    s1 = Get_AHour(0) '同1小时的
  End If
  s1 = Left(s1,32) :'Response.Write "<br>a."&s1&"<br>b."&s0&"<br>b."&mm&"<br>b."&m0
  If s0=s1 And mm-m0<=xMin Then
	Get_AHChk = true
  Else
	Get_AHChk = false
  End If
End Function 
'// 每1小时更换一次,n为超时分钟; 
'// Demo <input name="App28Code" type="hidden" value="<=Get_AHour(0)>" />
Function Get_AHour(xMin)
  Dim s,D,h,n
  D = DateAdd("n",(-1)*xMin,Now())
  h = DatePart("h",D) : If len(h)=1 Then h="0"&CStr(h)
  n = DatePart("n",D) : If len(n)=1 Then n="0"&CStr(n)
  s = MD5_Mem(Get_yyyymmdd(D)&h)&n ':Response.Write "<br>::::"&s
  Get_AHour = s
End Function 
' //每天更换一次
Function Get_AppDay()
  Dim s,d,n
  s = Get_hhmmss()
  If s>"234500" Or s<"001500" Then
    d = DateAdd("n",45,Now())
  Else
    d = Now()
  End If
  n = DatePart("d",d)
  Get_AppDay = n
End Function 


'/////////////////////////////////////////////////////////////////////////////////
'tApp25Code,tApp25Data:协议-附加码,协议-附加值 
'App25Code = "123_317_832_798_184"
Dim App25A,App25B,App25C,App25D
Sub App25Set()
 Dim s1,s2,s01,s02,n1,n2,a,i
 s1="" : s2=""
 s01 = "!$()*+,-/:;=?@^_`|~" '\/:*?""<>|=+-&#%
 s02 = "(@)!$:;=?^_`|~*+,-/" '(.){\'/}[]<>~!@
 n1 = Int(40/Len(s01)) '4000
 n2 = Int(40/Len(s02))
 a = Split(App25Code,"_")
 For i=0 To n1
   s1 = s1&s01&" "
   s2 = s2&s02&" "
 Next
 App25A = Trim(Mid(s1,Int(a(1))))
 App25B = Trim(Mid(s2,Int(a(2))))
 App25C = Trim(Mid(s2,Int(a(3))))
 App25D = Trim(Mid(s1,Int(a(4))))
End Sub



'/////////////////////////////////////////////////////////////////////////////////
'SwhCodRead = "N" 'Y/N 阅读协议-认证码



'/////////////////////////////////////////////////////////////////////////////////
'Dim App23RStr,App23RLen,App23RS01,App23RS02
'App23Check,App23Set
Sub App23Set()
  App23RStr = ""
  Randomize 
  App23RLen = Int(48*Rnd())+22 '4800,2200
  App23RLen = Int(App23RLen/2)
  For i=1 To App23RLen
	App23T1 = Int(88*Rnd())+11 '8888,1001
    App23RStr = App23RStr&App23T1
  Next
  App23RS01 = Left(App23RStr,4)*3.1415926 'Int(Mid(App23RStr,App23RLen*1,1)*3.1415926*Right(App23Date,1))
  App23RS02 = Right(App23RStr,4)*2.7182818 'Int(Mid(App23RStr,App23RLen*3,1)*2.7182818*Left(App23Date,1))
End Sub
'App23Date = "20100401" Len 2200~4800
Function App23Check(xFlag)
  Dim R23Flg,R23Chk,R23Str,r,f,n,S01,S02
  R23Flg = Request.Form(xFlag&"RFlg")
  R23Chk = Request.Form(xFlag&"RChk")
  R23Str = Request.Form(xFlag&"RStr")
  If R23Flg="" Or R23Chk="" Or R23Str="" Then
    r = false :Response.Write "23-1"
  ElseIf Len(R23Str)<20 Then '2000
    r = false ':Response.Write "23-2"
  ElseIf Len(R23Str) mod 2>0 Then
    r = false ':Response.Write "23-3"
  Else
    f = xFlag
	n = Mid(R23Flg,9)
	S01 = Left(R23Str,4)*3.1415926 'Int(Mid(R23Str,n*1,1)*3.1415926*Right(f,1))
    S02 = Right(R23Str,4)*2.7182818 'Int(Mid(R23Str,n*3,1)*2.7182818*Left(f,1))
	If S01&S02<>R23Chk Then
	  r = false ':Response.Write "23-4"
	ElseIf Len(R23Str) <> Int(n*2) Then
	  r = false ':Response.Write "23-5"
	Else
	  r = true ':Response.Write "23-6"
	End If
  End If
  'Response.End()
  App23Check = r
End Function
'Demo Code = /////////////////////////
'[mu_"&AppRand12&".asp] Dim App23RStr,App23RLen,App23RS01,App23RS02
'Call App23Set()
'<input type="hidden" name="<^=App23Date^>RFlg" id="<^=App23Date^>RFlg" value="<^=App23Date^><^=App23RLen^>" />
'<input type="hidden" name="<^=App23Date^>RChk" id="<^=App23Date^>RChk" value="<^=App23RS01^><^=App23RS02^>" />
'<input type="hidden" name="<^=App23Date^>RStr" id="<^=App23Date^>RStr" value="<^=App23RStr^>" />
'App23Flag = App23Check(App23Date) 
'If Not App23Flag Then msg = "错误R23：请重新提交 或 联系管理员!"
'/////////////////////////////////////////////////////////////////////////////////

%>