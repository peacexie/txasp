<%

Dim RsaT,RsaE,RsaD
RsaT=4592124
RsaE=7849
RsaD=4692169

Function RsaGetPara()
Dim RsaP,RsaQ,RsaN
  RsaP = RsaGetPQ(1201,2401)
  RsaQ = RsaGetPQ(2401,3601)
  RsaN = RsaP * RsaQ
  RsaT = (RsaP-1)*(RsaQ-1)
  RsaE = RsaGetE(1234,9876) 
  RsaD = RsaGetD() 
Response.Write "<br>P-Q-N:"&RsaP&"-"&RsaQ&"-"&RsaN
Response.Write "<br>T-E-D:"&RsaT&"-"&RsaE&"-"&RsaD
Response.Write "<br>RsaT="&RsaT&"<br>RsaE="&RsaE&"<br>RsaD="&RsaD
End Function

Function RsaGetPQ(xMin,xMax)
Dim i,xn : xn = PNum9999(xMin,xMax)
  For i = xn To 65535 Step 2 
    If PNumFlag(i) = "YES" Then
	 RsaGetPQ = i 
	 Exit For
    End If
  Next
End Function
Function RsaGetE(xMin,xMax)
Dim i,j : j=PNum9999(xMin,xMax)
  For i = j To 65535 Step 2  
    If PNumPair(i,RsaT) = "YES" Then
	  RsaGetE = i
	  Exit For
	End If
  Next
End Function 
Function RsaGetD()
Dim j,i,dx,d
  For i = 1001 To 9999
	dx = i * RsaT + 1
	If (dx/RsaE)=Int(dx/RsaE) Then
	  RsaGetD = dx/RsaE
	  Exit For
	End If
  Next
End Function 

Function PNum9999(xMin,xMax)
Dim xn : Randomize
  xn = (xMax - xMin) * Rnd + xMin  
  PNum9999 = Int(xn/2)*2+1 ' 单数
End Function 
Function PNumFlag(xN)
Dim N,Flag,i
N = xN : Flag = "YES"
  If (N<2) Or (N>65535) Then
    Flag = "NO"
  ElseIf N = 2 Then
    Flag = "YES"
  ElseIf (N Mod 2) = 0 Then
    Flag = "NO"
  Else
    For i = 3 To Sqr(N) Step 2
	  If N Mod i = 0 Then
	    Flag = "NO"
		Exit For
	  End If
	Next
  End If
PNumFlag = Flag
End Function 
Function PNumPair(xN,xM)
Dim t,Flag,i,N,M
N = xN : M = xM : Flag = "YES"
  If N > M Then
    t = N : N = M : N = t
  End If
  If (M>65535) Then M = 65535
  If N < 1 Then Flag = "NO"
  If M Mod N = 0 Then
    Flag = "NO"
  Else
    For i = 2 To N/2 
	  If ((N Mod i)=0) AND ((M Mod i)=0) Then
	    Flag = "NO"
		Exit For
	  End If
	Next    
  End If
PNumPair = Flag
End Function 

Function Rsa_Enc(xStr)
Dim msg,i,n2,c2,str
msg = xStr : str = ""
For i = 1 To Len(msg)
  n2 = Asc(Mid(msg,i,1))
  c2 = (n2*RsaE)
  c2 = c2 - RsaT*Int(c2/RsaT)
  c2 = Hex(c2)
  c2 = Right("000000"&c2,6)
  str = str&c2
Next
Rsa_Enc = str
End Function
Function Rsa_Dec(xStr)
Dim msg,oStr,i,j,s1,n1,n2,ch,na,ns
msg = xStr : oStr = ""
For i = 1 To Len(msg) Step 6
  s1 = Mid(msg,i,6) : n1 = 0
  For j = 6 To 1 Step -1
    ch = Mid(s1,j,1) 
	na = Asc(ch) 
	If     ch&"" >= "A" And ch&"" <= "F" Then
	  na = na - 55
	ElseIf ch&"" >= "0" And ch&"" <= "9" Then
	  na = na - 48      
	End If
	n1 = n1 + na * 16^(6-j)
  Next
  n2 = (n1*RsaD) 
  n2 = n2 - RsaT*Int(n2/RsaT)
  oStr = oStr & Chr(n2)
Next
Rsa_Dec = oStr
End Function

Function Rsa_pEnc(xPW,xID) '加密 Rsa_pEnc(Pub_Confg&UsrPW&UsrID)
Dim lID,lPW,tStr,tLen
 lID=Len(xID)  : lPW=Len(xPW)
 tLen=lID+lPW : tStr=Left(xPW&xID,36)
 If tLen Mod 2=1 Then tLen=tLen+1 ' tLen为单数
 tStr = RSA_Enc(tStr)
 lStr=Left(tStr,tLen-2)
 rStr=Right(tStr,tLen+2)
 mStr=Mid(tStr,tLen-1,Len(tStr)-2*tLen)
 tStr=mStr&rStr&lStr
Rsa_pEnc = tStr
End Function

Function Rsa_pDec(xStr,xID) '还原
Dim lID,lPW,tStr,tLen
 lID=Len(xID) 
 tLen=Len(xStr)/6
 lPW=tLen-Len(lID)
 If tLen Mod 2=1 Then tLen=tLen+1
 mStr=Left(xStr,Len(xStr)-2*tLen)
 lStr=Right(xStr,tLen-2)
 rStr=Mid(tStr,Len(xStr)-2*tLen+1,tLen+2)
 tStr=lStr&mStr&rStr
 tStr=RSA_Dec(tStr)
 tStr=Left(tStr,tLen-lID)
Rsa_pDec = tStr
End Function


%>
