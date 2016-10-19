
<%

Const Pub_IDStr = "0123456789ABCDEFGHJKMNPQRSTUVWXY" ' Ii  Ll  Oo  Zz
'Const Pub_IDEnc = "V6BCAQHNUEW4GRP8T570FYS1XMD93JK2" 'App31Code

'Response.Write "<br>------6时间---2版本-2TID-6标识地点--"
'Response.Write "<br>12345678-1234-1234-1234-123456789012"
'Response.Write "<br>123456789012345678901234567890123456"
'Response.Write "<br>"&Get_GUID("99.99","IP")&" --- myGUID"
Function Get_GUID(sVer,sAddr)
  Dim str1,str2,str3,str4
  str1 = Get_AutoID(12)
  str1 = Left(str1,8)&"-"&Right(str1,4)
  str2 = Get_GUVer(sVer)
  str3 = Base_32(Session.SessionID(),16,4,"Right")
  str4 = Get_GUAddr(sAddr)
  Get_GUID = str1&"-"&str2&"-"&str3&"-"&str4
End Function 
Function Get_GUVer(sVer)
  Dim str : str = sVer
  If inStr(str,".")>0 Then 'Max:655.35; 654.99
    a = Split(str,".")
	str = a(1) : if(Len(a(1))=1) Then a(1)="0"&a(1)
	str = a(0)&a(1) 'Replace(str,".","")
  End If
  'Response.Write str
  If isNumeric(str) Then
    num = CLng(str)
	str = Hex(num)
	If Len(str)<4 Then
	  str = Right("0000"&str,4)
	Else
	  str = Left(str,4)
	End If
  Else
    str = Rnd_ID("",4)
  End If
  Get_GUVer = str
End Function 
Function Get_GUAddr(sAddr)
  Dim s,a,i,s0,s1,s2
  If sAddr="IP" Then
    s1 = Get_CIP()
	a = Split(s1,".")
	For i=0 To 3 
	  s0 = Hex(a(i)) : If Len(s0)=1 Then s0="0"&s0
	  s = s&s0
	Next
  ElseIf sAddr="Mac" Then
    s = ""
  ElseIf sAddr="Code" Then
    Const c1 = "0086" : Const c2 = "0769" 
	s1 = Hex(c1) : If Len(s1)=1 Then s1="0"&s1
	s2 = Hex(c2) : If Len(s2)=1 Then s2="0"&s2
	s = s1&s2
  ElseIf Len(sAddr)>7 Then
    s = sAddr
  Else
    s = Rnd_ID("",12)
  End If
  If Len(s)<12 Then s=s&Rnd_ID("",12-Len(s))
  If Len(s)>12 Then s=Left(s,12)
  Get_GUAddr = s
End Function  

'eID = Get_IDUnc(rID,0)
Function Get_IDUnc(xStr,xPos)
  Dim i,c,p,s :s=""
  For i=1 To Len(xStr)
    c = Mid(xStr,i,1)
	p = inStr(App31Code,c)
	If p>0 Then
	  p = p-xPos :If p<1 Then p=p+Len(Pub_IDStr)
	  c = Mid(Pub_IDStr,p,1)
	  s = s&c
	Else
	  s = s&c
	End If
  Next
  Get_IDUnc = s
End Function 
'eID = Get_IDEnc(rID,0)
Function Get_IDEnc(xStr,xPos)
  Dim i,c,p,s :s=""
  For i=1 To Len(xStr)
    c = Mid(xStr,i,1)
	p = inStr(Pub_IDStr,c)
	If p>0 Then
	  p=p+xPos
	  If p>Len(App31Code) Then p=p-Len(App31Code)
	  s = s&Mid(App31Code,p,1)
	Else
	  s = s&c
	End If
  Next
  Get_IDEnc = s
End Function 
'rID = Get_IDRnd()
Function Get_IDRnd()
  Dim i,t,s,a(31),b(31) '0~32
  Randomize :s=""
  For i=0 To 31
    b(i) = Mid(Pub_IDStr,i+1,1)
  Next
  For i=0 To 31
    a(i) = Int(32*Rnd()) '(rMax-1)
	t = b(a(i))
	b(a(i)) = b(i)
	b(i) = t
  Next
  For i=0 To 31
    s = s&b(i)
  Next
  Get_IDRnd = s
End Function 


Function Get_yyyymmdd(xDate)
  Dim ny,nm,nd,mm,dd,tDate
  If isDate(xDate) Then
    tDate = xDate
  Elseif isNumeric(xDate) Then
    tDate = DateAdd("d",xDate,Now())
  Else
    tDate = Now()
  End If
   ny = DatePart("yyyy",tDate)
   nm = DatePart("m",tDate)
   nd = DatePart("d",tDate)
   If Len(nm) = 1 Then
      mm = "0" & CStr(nm)
   Else
      mm = CStr(nm)
   End If
   If Len(nd) = 1 Then
      dd = "0" & CStr(nd)
   Else
      dd = CStr(nd)
   End If
   Get_yyyymmdd = ny & mm & dd
End Function   

Function Get_hhmmss()
  Dim hh,mm,ss,xhh,xmm,xss
   hh = DatePart("h",Now())
   mm = DatePart("n",Now())
   ss = DatePart("s",Now())
   If len(hh) = 1 Then
      xhh = "0" & CStr(hh)
   Else
      xhh = CStr(hh)
   End if
   If len(mm) = 1 Then
      xmm = "0" & CStr(mm)
   Else
      xmm = CStr(mm)
   End if
   If len(ss) = 1 Then
      xss = "0" & CStr(ss)
   Else
      xss = CStr(ss)
   End if
   Get_hhmmss = xhh & xmm & xss
End Function   

Function Get_mSec()
Dim s,n 
   n = Timer() '//86400.00
   s = Int(1000*(n-int(n)))
Get_mSec = Right("00"&s,3)
End Function 

Function Fmt_Time(xStr,xType)
  Dim s1,s2,x1,x2,obj 'xType=D,T
   s1 = Left(xStr,8)
   s2 = Mid(xStr,9,6)
   s3 = Mid(xStr,15,3)
   x1 = Left(s1,4)& "-" &Left(Right(s1,4),2)& "-" &Right(s1,2) ' YYYY-MM-DD
   x2 = Left(s2,2)& ":" &Left(Right(s2,4),2)& ":" &Right(s2,2) ' HH:MM:SS
   If xType = "D" or xType = "d" Then ' YYYY-MM-DD
     obj = x1
   ElseIf xType = "T" or xType = "t" Then ' HH:MM:SS
     obj = x2
   ElseIf xType = "F" or xType = "f" Then ' YYYY-MM-DD HH:MM:SS
     obj = x1&" "&x2
   ElseIf xType = "X" or xType = "x" Then ' YYYY-MM-DD HH:MM:SS.999
     obj = x1&" "&x2&"."&s3
   Else
     obj = x1
   End If
   Fmt_Time = obj
End Function 


Function Get_9999ID(xStr,xLen)
  Dim obj,i
   obj = ""  
   for i = 1 to ( xLen - Len(xStr) )
      obj = obj & "0"
   next
    Get_9999ID = obj & xStr
End Function

'YYYY,YY,MM,DD,MD
'HNSX,R8,S8
Function Get_FmtID(xFmt,xTime)
  Dim sTim,sSec,sTmp,y4,mm,dd,md,hnsx,pn 
  sTmp = xFmt
  If xTime="" Then 
    sTim = Now()
	sSec = Timer()*10
  Else
	sTim = xTime :Dim th,tm,ts
    th = DatePart("h",sTim)*3600*10
    tm = DatePart("n",sTim)*60*10
    ts = DatePart("s",sTim)*10
	sSec = th + tm + ts
  End If
  y4 = Year(sTim) ' ------ yyyy,yy
  If inStr(sTmp,"yyyy")>0 Then
    sTmp = Replace(sTmp,"yyyy",y4)
  End If
  If inStr(sTmp,"yy")>0 Then
	sTmp = Replace(sTmp,"yy",Right(y4,2))
  End If
  mm = DatePart("m",sTim) ' ------ mm,dd,md 
  dd = DatePart("d",sTim) 
  If inStr(sTmp,"mm")>0 Then
	If len(mm)=1 Then mm="0"&CStr(mm)
	sTmp = Replace(sTmp,"mm",mm)
  End If
  If inStr(sTmp,"dd")>0 Then
	If len(dd)=1 Then mm="0"&CStr(dd)
	sTmp = Replace(sTmp,"dd",dd)
  End If
  If inStr(sTmp,"md")>0 Then
	md = Mid(Pub_IDStr,mm+1,1)&Mid(Pub_IDStr,dd+1,1)
	sTmp = Replace(sTmp,"md",md)
  End If
  If inStr(sTmp,"hnsx")>0 Then ' ------ hnsx
	hnsx = Base_32(sSec,32,4,"") '1/10s 'Timer() '//86400.00 
	sTmp = Replace(sTmp,"hnsx",hnsx)
  End If
  If inStr(sTmp,"r")>0 Then ' ------ rn
    pn = Mid(sTmp,inStr(sTmp,"r")+1,1) ':echo "---"&pn
	sTmp = Replace(sTmp,"r"&cStr(pn),Rnd_ID("KEY",pn))
  End If
  If inStr(sTmp,"s")>0 Then ' ------ sn
    pn = Mid(sTmp,inStr(sTmp,"s")+1,1) ':echo pn
	sTmp = Replace(sTmp,"s"&cStr(pn),Base_32(Session.SessionID(),32,pn,"Right"))
  End If
  Get_FmtID = sTmp
End Function

Function Get_AutoID(xN)
Dim YMD,HMS,str,tDate,ny,nm,nd,RDX
   tDate = Now()
   ny = DatePart("yyyy",tDate)
   nm = DatePart("m",tDate)
   nd = DatePart("d",tDate)
  YMD = Hex(ny*380+nm*31+nd)
  HMS = Hex(Timer()*100)
  YMD = Right("00000"&YMD,6)
  HMS = Right("00000"&HMS,6)
  str = YMD&Left(HMS,2)&"-"&Right(HMS,4)
  If xN=12 Then 
    Get_AutoID = YMD&HMS
  ElseIf xN=15 Then 
    Get_AutoID = str&Rnd_ID("KEY",2)
  ElseIf xN=18 Then 
    Get_AutoID = str&Base_32(Session.SessionID(),16,3,"Right")&Rnd_ID("KEY",2)
  Else
   If xN>=20 Then
   	Get_AutoID = str&"-"&Base_32(Session.SessionID(),16,4,"Right")&"-"&Rnd_ID("KEY",(xN-19))
   Else
   	Get_AutoID = Left(str&"-"&Rnd_ID("KEY",6),xN)
   End If
  End If
End Function

Function Rnd_ID(xType,xLen2)
  Dim objStr,orgNum,orgCap,orgLow,orgSp1,orgSp2,oStr,rStr,rMax,i,nx
  orgNum = "0123456789"                   ' 10   xType = 0,KEY,0aA
  'orgCap = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"   ' 26  
  'orgLow = "abcdefghijklmnopqrstuvwxyz"   ' 26   .  -  _  $  |  !  #  [  ]
  orgKey = Mid(Pub_IDStr,11)              ' 22   Ii  Ll  Oo  Zz
  org16B = "0123456789ABCDEF" 
  oStr = "" :rStr = ""  
  If xType = "0" Then
	rMax = 10 :oStr = orgNum
  ElseIf uCase(xType) = "A" Then
	rMax = 22 :oStr = orgKey
  ElseIf uCase(xType) = "KEY" Then
	rMax = 32 :oStr = Pub_IDStr '10+22
  Else 'If xType = "0aA" Then
	rMax = 16 :oStr = org16B'16
  End If 
      Randomize :objStr = ""
	  if bLen2 Mod 2=0 then bLen2=bLen2+1
	  For i=1 To xLen2
	    xn = Int(rMax*Rnd()) '+ (xMax-1)
	    objStr = objStr&Mid(oStr,xn+1,1) 
	  Next  
  Rnd_ID = Left(objStr,xLen2) 
End Function
Function Rnd_Base(xBase,xLen)
  Dim xMax,sObj,xn,i
  If xBase="" Then xBase=Pub_IDStr
  Randomize
  xMax = Len(xBase) :sObj = ""
  For i=1 To xLen
    xn = Int(xMax*Rnd()) '+ (xMax-1)
	sObj = sObj&Mid(xBase,xn+1,1)
  Next  
  Rnd_Base = sObj
End Function


'NextID(0024,0012,4)=0028 ::: MZ98 - S110028
Function Next_ID(xOld,xMin,xStep)
Dim nLen,t1,tf,tn,hf,OldID,i,c,n,s
  nLen = Len(xMin) : s = ""
  t1 = "0123456789"
  tf = "Num" : tn = t1 : hf = 0 '(进位) 
  If Len(xOld)<nLen Then
    Next_ID = xMin
  Else
    OldID = Left(xOld,nLen)
	For i = nLen To 1 Step -1
	    c = Mid(OldID,i,1)
	    If tf<>"Num" Or c>"9" Or i=1 Then tf="Chr"
	    If tf="Chr" Or Len(xOld)<4 Then tn = Pub_IDStr
	 n = inStr(tn,c) ' 位置 1~N
	 If n>0 Then
	    n = n + Int(hf)
	    If i=nLen Then n = n + Int(xStep)
	    If n > Len(tn) Then
	      n = n-Len(tn)
		  hf = 1
	    Else
		  hf = 0
	    End If
		c = Mid(tn,n,1)
	 Else
	    c = c 'ILOZ不在列表中则
	 End If
	    s = c&s
	Next
    If hf = 1 Then
	  s = "_"&s
	End If
	Next_ID = s
  End If 
End Function
'Response.Write inStr("abc","a")

'Demo:Session.SessionID(),32,3,"Right"
Function Base_32(xNum,xBase,xLen,xDir)
  Dim n,m0,m1,i,ni,si,s0 '864000(10)=QBO0(32) 'xDir = Left,Right,
  n = xNum : s0 = ""
  m0 = xBase^xLen 'max
  'Response.Write "<br>1.0:"&m0
  If n>=m0 Then
    If xDir="Right" Then 
	  m1 = Right(n,Len(m0))
	  If m1>m0 Then m1=Right(n,Len(m1)-1)
	Else
	  m1 = Left(n,Len(m0))
	  If m1>m0 Then m1=Left(n,Len(m1)-1)
	End If 
    n = m1
  End If
  'Response.Write "<br>2.0:"&n
  For i=xLen-1 To 0 Step -1
    If i>0 Then
      ni = Int(n/(xBase^i))
    Else
      ni = n
    End If
	'Response.Write "<br>3.0:"&ni
	si = Mid(Pub_IDStr,ni+1,1)
	s0 = s0&si
	n = n Mod (xBase^i)
  Next
  Base_32 = s0
End Function


%>

