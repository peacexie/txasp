<%
'--- UNICODE,UTF-8规则 ------------------------------------------------------
'UNICODE  - UTF-8 
'0000 - 007F 0xxxxxxx 
'0080 - 07FF 110xxxxx 10xxxxxx 
'0800 - FFFF 1110xxxx 10xxxxxx 10xxxxxx 
'       8F6C 11101000 10111101 10101100
'--- UTF-8>UNICODE ------------------------------------------------------
'iNum:16进制; 返回10进制 
'Server.URLEncode("转")=E8BDAC '%E8%BD%AC
'cNum_U8toU(8F6C)=36716
'ChrW(36716)="转"
Function cNum_U8toU(iHex) 
  Dim sBin,tBin
  sBin = c16to2(iHex) ':Response.Write sBin
  If Left(sBin,1)="0" Then
    tBin = Mid(sBin,2)
  ElseIf Left(sBin,3)="110" Then
    tBin = Mid(sBin,4,5)&Mid(sBin,11,6)
  ElseIf Left(sBin,4)="1110" Then
    tBin = Mid(sBin,5,4)&Mid(sBin,11,6)&Mid(sBin,19,6)
  End If
  'Response.Write tBin
  cNum_U8toU = c2to10(tBin) 'c2to16(tBin) '
End Function
'--- UNICODE>UTF-8 ------------------------------------------------------
'iNum:10进制; 返回16进制 
'AscW("转")=36716
'cNum_UtoU8(36716)=8F6C
Function cNum_UtoU8(iNum) 
  Dim sResult,jNum,b,c
  jNum = iNum
  If jNum = "" Then
    Exit Function
  End If
  sResult = ""
  If (jNum < 128) Then
    sResult = Hex(jNum)
  ElseIf (jNum < 2048) Then
    sResult = Hex(&H80 + (jNum And &H3F))
    jNum = jNum \ &H40
    b = Hex(&HC0 + (jNum And &H1F))
    sResult =  b& sResult
  ElseIf (jNum < 65536) Then
    sResult = Hex(&H80 + (jNum And &H3F)) ':Response.Write "<br>3."&a
    jNum = jNum \ &H40
    b = Hex(&H80 + (jNum And &H3F)) ':Response.Write "<br>3."&b
    sResult = b& sResult
    jNum = jNum \ &H40
    c = Hex(&HE0 + (jNum And &HF))  ':Response.Write "<br>3."&c
    sResult = c& sResult
  End If
  cNum_UtoU8 = sResult
End Function

'=====================================
' Unicode: Unicode编码, &#x54;&#x65;&#x73;&#x74;&#x6D4B;&#x8BD5;
'=====================================
Function conv_Unicode(xStr)
   Dim i,ch,s
   For i=1 to len(xStr)
	  ch=Mid(xStr,i,1)
	  cx=Hex(AscW(ch))
	  s=s&"&#x"&cx&";"
   Next
   conv_Unicode = s
End Function


'=====================================
' cUrl_Encode: 改变当前会话编码格式，编码后还原
'=====================================
Function cUrl_Encode(xStr,xSet,xNow)
  Session.Codepage = xSet
  cUrl_Encode = Server.URLEncode(xStr)
  Session.Codepage = xNow'65001'还原当前会话编码
End Function

'=====================================
' Url 编码,解码
'=====================================
Function cUrl_gb2u8(szInput)   
  Dim wch, uch, szRet,x   
  Dim nAsc, nAsc2, nAsc3   
  If szInput = "" Then '如果输入参数为空，则退出函数    
	cUrl_gb2u8 = szInput   
	Exit Function  
  End If  
  For x = 1 To Len(szInput) '开始转换  
	wch = Mid(szInput, x, 1)   
	nAsc = AscW(wch)   
	If nAsc < 0 Then nAsc = nAsc + 65536   
	If (nAsc And &HFF80) = 0 Then  
	  szRet = szRet & wch   
	Else  
	  If (nAsc And &HF000) = 0 Then  
	    uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)   
	    szRet = szRet & uch   
	  Else  
	    uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _   
	    Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _   
	    Hex(nAsc And &H3F Or &H80)   
	    szRet = szRet & uch   
	  End If  
	End If  
  Next  
  cUrl_gb2u8= szRet   
End Function 
Public Function cUrl_Decode(xStr)
  Dim s,i,l,c,t,n : s="" : l=Len(xStr)
  For i=1 To l
	c=Mid(xStr,i,1)
	If c<>"%" Then
	  s = s & c
	Else
	  c=Mid(xStr,i+1,2) : i=i+2 : t=CInt("&H" & c)
	  If t<&H80 Then
		s=s & Chr(t)
	  Else
		c=Mid(xStr,i+1,3)
		If Left(c,1)<>"%" Then
		  cUrl_Decode=s
		  Exit Function
		Else
		  c=Right(c,2) : n=CInt("&H" & c)
		  t=t*256+n-65536
		  s = s & Chr(t) : i=i+3
		End If
	  End If
	End If
  Next
  cUrl_Decode=s
End Function
Function cUrl_Escape(str) 
    dim i,s,c,a 
    s="" 
    For i=1 to Len(str) 
        c=Mid(str,i,1) 
        a=ASCW(c) 
        If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then 
            s = s & c 
        ElseIf InStr("@*_+-./",c)>0 Then 
            s = s & c 
        ElseIf a>0 and a<16 Then 
            s = s & "%0" & Hex(a) 
        ElseIf a>=16 and a<256 Then 
            s = s & "%" & Hex(a) 
        Else 
            s = s & "%u" & Hex(a) 
        End If 
    Next 
    cUrl_Escape = s 
End Function
Function cUrl_UnEsc(str) 
    dim i,s,c 
    s="" 
    For i=1 to Len(str) 
        c=Mid(str,i,1) 
        If Mid(str,i,2)="%u" and i<=Len(str)-5 Then 
            If IsNumeric("&H" & Mid(str,i+2,4)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4))) 
                i = i+5 
            Else 
                s = s & c 
            End If 
        ElseIf c="%" and i<=Len(str)-2 Then 
            If IsNumeric("&H" & Mid(str,i+1,2)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+1,2))) 
                i = i+2 
            Else 
                s = s & c 
            End If 
        Else 
            s = s & c 
        End If 
    Next 
    cUrl_UnEsc = s 
End Function 

'=====================================
'用途:將UTF-8編碼漢字轉換為GB2312碼，兼容英文和數字["岁月联盟"提供] 
'用法:Response.write cUrl_ut2gb("%E9%83%BD%E5%B8%82%E6%83%85%E7%B7%A3 %E6%98%9F%E5%BA%A7") 
'=====================================
function cUrl_ut2gb(UTFStr) 
for Dig=1 to len(UTFStr) 
  if mid(UTFStr,Dig,1)="%" then 
    if len(UTFStr) >= Dig+8 then 
      GBStr=GBStr & cUrl_Chinese(mid(UTFStr,Dig,9)) 
      Dig=Dig+8 
    else 
      GBStr=GBStr & mid(UTFStr,Dig,1) 
    end if 
  else 
    GBStr=GBStr & mid(UTFStr,Dig,1) 
  end if 
next 
cUrl_ut2gb=GBStr 
end function 
function cUrl_Chinese(x) 
  A=split(mid(x,2),"%") 
  i=0 
  j=0 
  for i=0 to ubound(A) 
    A(i)=c16to2(A(i)) 
  next 
  for i=0 to ubound(A)-1 
    DigS=instr(A(i),"0") 
    Unicode="" 
    for j=1 to DigS-1 
      if j=1 then 
        A(i)=right(A(i),len(A(i))-DigS) 
        Unicode=Unicode & A(i) 
      else 
        i=i+1 
        A(i)=right(A(i),len(A(i))-2) 
        Unicode=Unicode & A(i) 
      end if 
    next 
    if len(c2to16(Unicode))=4 then 
      cUrl_Chinese=cUrl_Chinese & chrw(int("&H" & c2to16(Unicode))) 
    else 
      cUrl_Chinese=cUrl_Chinese & chr(int("&H" & c2to16(Unicode))) 
    end if 
  next 
end function 


'=====================================
' Convert Str,Bin
'=====================================
function cBin_s2b(str, cSet) 
''''''''''''''''''''''''''''' 
' 字符串转二进制函数 By shawl.qiu 
'   http://blog.csdn.net/btbtd 
'''''''''''''''' 
' 参数说明 
'''''''''' 
' charSet: 字符串默认编码集, 如不指定, 则默认为 gb2312 
'''''''''' 
' sample call: response.binaryWrite cBin_s2b(str, "utf-8") 
''''''''''''''''''''''''''''' 
	dim stm  
	set stm = CreateObject("adodb.stream") 
		with stm 
			.type = 2  
			if cSet<>"" then 
				.charSet = cSet 
			else 
				.charSet = "gb2312" 
			end if 
			.open 
			.writeText str 
			.Position = 0 
			.type = 1 
			cBin_s2b = .Read 
			.close 
		end with 'shawl.qiu code' 
	set stm = nothing 
end function 
 
function cBin_b2s(str, cSet) 
''''''''''''''''''''''''''''' 
' 二进制转字符串函数 By shawl.qiu 
'   http://blog.csdn.net/btbtd 
'''''''''''''''' 
' 参数说明 
'''''''''' 
' charSet: 字符串默认编码集, 如不指定, 则默认为 gb2312 
'''''''''' 
' sample call: response.write cBin_b2s(midB(cBin_s2b(str, "utf-8"),1),"utf-8") 
''''''''''''''''''''''''''''' 
' 注意: 二进制字符串必须先用 midB(binaryString,1) 读取(可自定读取长度). 
''''''''''''''''''''''''''''' 
	dim stm  
	set stm = CreateObject("adodb.stream") 
		with stm
			.type = 2  
			.open 
			.writeText str 
			.Position = 0 
			if cSet<>"" then 
				.CharSet = cSet 
			else  
				.CharSet = "gb2312" 
			end if 
				cBin_b2s = .ReadText 
			.Close 
		end with 'shawl.qiu code' 
	set stm = Nothing 
end function 


'=====================================
' Convert 2,10,16
'=====================================
function c2to16(x) 
  i=1 
  for i=1 to len(x) step 4 
    c2to16=c2to16 & hex(c2to10(mid(x,i,4))) 
  next 
end function 

function c2to10(x) 
  c2to10=0 
  if x="0" then exit function 
  i=0 
  for i= 0 to len(x) -1 
    if mid(x,len(x)-i,1)="1" then c2to10=c2to10+2^(i) 
  next 
end function 

function c16to2(x) 
  i=0 
  for i=1 to len(trim(x)) 
    tempstr= c10to2(cint(int("&h" & mid(x,i,1)))) 
    do while len(tempstr)<4 
      tempstr="0" & tempstr 
    loop 
    c16to2=c16to2 & tempstr 
  next 
end function 

function c10to2(x) 
  mysign=sgn(x) 
  x=abs(x) 
  DigS=1 
  do 
    if x<2^DigS then 
      exit do 
    else 
      DigS=DigS+1 
    end if 
  loop 
  tempnum=x 
  i=0 
  for i=DigS to 1 step-1 
    if tempnum>=2^(i-1) then 
      tempnum=tempnum-2^(i-1) 
      c10to2=c10to2 & "1" 
    else 
      c10to2=c10to2 & "0" 
    end if 
  next 
  if mysign=-1 then c10to2="-" & c10to2 
end function


'=====================================
' Convert Else,Sniff
'=====================================
Function conv_Sniff( sData )
	Dim oRE
	Set oRE = New RegExp
	oRE.IgnoreCase	= True
	oRE.Global		= True
	Dim aPatterns
	aPatterns = Array( "<!DOCTYPE\W*X?HTML", "<(body|head|html|img|pre|script|table|title)", "type\s*=\s*[\'""]?\s*(?:\w*/)?(?:ecma|java)", "(?:href|src|data)\s*=\s*[\'""]?\s*(?:ecma|java)script:", "url\s*\(\s*[\'""]?\s*(?:ecma|java)script:" )
	Dim i
	For i = 0 to UBound( aPatterns )
		oRE.Pattern = aPatterns( i )
		If oRE.Test( sData ) Then
			conv_Sniff = True
			Exit Function
		End If
	Next
	conv_Sniff = False
End Function





%>