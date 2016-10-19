<%

'说明: 此文件函数，为公共函数文件，
'所有 变量/函数 以smp开头

Dim smpSoapHead,smpSoapFoot
smpSoapHead = ""&_
			 "<?xml version=""1.0"" encoding=""utf-8""?>"&_
			 "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"&_
			 "<soap:Body>"
smpSoapFoot = "</soap:Body></soap:Envelope>"
			 

' 格式化号码 时间yyyymmddhhnnss
Function smpFmtTime(xTime)
  Dim ny,nm,nd,hh,mm,ss
  If isDate(xTime) Then
   ny = DatePart("yyyy",xTime)
   nm = DatePart("m",xTime)
   nd = DatePart("d",xTime)'&"-"
   If Len(nm) = 1 Then nm = "0" & CStr(nm)
   If Len(nd) = 1 Then nd = "0" & CStr(nd)
   hh = DatePart("h",xTime)
   mm = DatePart("n",xTime)
   ss = DatePart("s",xTime)
   If len(hh) = 1 Then hh = "0" & CStr(hh)
   If len(mm) = 1 Then mm = "0" & CStr(mm)
   If len(ss) = 1 Then ss = "0" & CStr(ss)
   smpFmtTime = ny & nm & nd & hh & mm & ss
  Else
    smpFmtTime = ""
  End If
End Function 
' 格式化号码:(0769)12345678; 0769-1234-5678
Function smpFmtTNum(xStr) 
  Dim s : s = xStr 
  s = Show_RExp(s,"\[[^>]*?\]","") '过滤姓名: 198-1234-5678 [Xie永顺]
  s = Replace(Replace(s,"<",""),">","") ' 去掉 <>
  s = Replace(Replace(Replace(s,"(",""),")",""),"-","") ' (-)
  s = Replace(Replace(Replace(s,vbcrlf,";"),vbcr,";"),vblf,";") ' 回车换行
  s = Replace(Replace(s," ",""),"　",";") ' 去掉 空格
  s = Replace(Replace(s,",",";"),"，",";") ' 去掉 ,
  s = Replace(Replace(s,";;;",";"),";;",";") ' 去掉 连续空格
  smpFmtTNum = s
End Function
' 判断是否合法 号码
Function smpChkTGrp(xStr) 
  Dim s : s = xStr
  If Len(s)<10 Then 'Or Len(s)>15 
    smpChkTGrp = false 'echo("1")
  Else
    s = smpFmtTNum(s)
    If Not isNumeric(s) then 
      smpChkTGrp = false 'echo("2")
    Else
      smpChkTGrp = true 'echo("3")
    End If
  End If
End Function
' 判断是否合法 号码
Function smpChkTNum(xStr) 
  Dim s : s = xStr
  If Len(s)<10 Or Len(s)>15 Then
    smpChkTNum = false
  ElseIf Not isNumeric(s) then 
    smpChkTNum = false
  Else
    smpChkTNum = true
  End If
End Function
' 内容分割条数
Function smpCntCont(xCont) 
  Dim n :n = Len(xCont)
  If n<=70 Then
    n = 1
  Else
   if n mod 65=0 then
	n = int(n/65)
   else
	n = int(n/65)+1
   end if
  End If
  smpCntCont = n
End Function
' 显示内容
Function smpShowCont(xStr) ' Save,Edit,UBB,show text in html
   Dim xText : xText=xStr
   xText = xText&""
   xText = Replace(xText, "<", "&lt;")
   xText = Replace(xText, ">", "&gt;")
   smpShowCont = xText
End Function 
  

'// 创建对象
Sub smpOSet() 
  If smcMClass<>"" Then
    Set smcMObj = Server.CreateObject(smcMClass)
  End If
End Sub
'// 注销对象
Sub smpOEnd() 
  Set smcMObj = Nothing
End Sub
'// 发送 Soap对象，返回状态数组
Function smpHttp(xUrl,xHost,xSoap,xAct)
  Dim oHttp,sRequest,bHttp(2)
  Set oHttp = Server.CreateObject("Msxml2.xmlHttp") 
	Dim p,t,sRndID,sSoap 
	sRndID = Rnd_ID("",12) :p = inStr(xSoap,">") :t = Left(xSoap,p) 
	sSoap = Replace(xSoap,t,t&"<Peace_"&sRndID&"_RndID>"&Timer()&"</Peace_"&sRndID&"_RndID>")
	sRequest = smpSoapHead&sSoap&smpSoapFoot
	'Response.Write "<br>"&smpXmlShow(sRequest)&"<br>"
	oHttp.Open "POST",xUrl,false
	oHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
	oHttp.setRequestHeader "HOST",xHost
	oHttp.setRequestHeader "Content-Length",Len(sRequest)
	oHttp.setRequestHeader "SOAPAction", "http://tempuri.org/"&xAct&"" 'tempuri.org,sdkhttp.eucp.b2m.cn
	oHttp.Send(sRequest)
	bHttp(0) = oHttp.Status
	bHttp(1) = oHttp.StatusText
	bHttp(2) = oHttp.ResponseText 
	smpHttp = bHttp
  Set oHttp = Nothing
End Function 
'// 发送 Http对象，返回状态数组
Function smpPost(xUrl,xAct,xData)
  Dim oHttp,bHttp(2),sPost,sUrl
  sUrl = xUrl&xAct&".action" 'for Test 48000
  Dim sRndID :sRndID = "Peace_"&Rnd_ID("",12)&"_RndID="&Timer()&"&"
  sPost = Replace(smsXmlUser&xData,"?","?"&sRndID&"")
  'Response.Write "<br><br>Len(sPost):"&Len(sPost)
  'Response.Write "<br>"&sUrl&sPost&"<br>"
  Dim fSend :fSend = "xGET" 'GET,POST
  Set oHttp = Server.CreateObject("Msxml2.xmlHttp")
  With oHttp 
	If fSend="GET" Then
	  .Open "GET",sUrl&sPost,False," ", " " 
	  .Send()
	Else
	  .Open "POST",sUrl,False," ", " " 
	  .setRequestHeader "Content-Length",Len(sPost) 
	  .setRequestHeader "Content-Type","application/x-www-form-urlencoded" 
	  .Send(sPost)	
	End If
	bHttp(0) = oHttp.Status
	bHttp(1) = oHttp.StatusText
	bHttp(2) = oHttp.ResponseText 
  End With 
  smpPost = bHttp
  Set oHttp = Nothing
End Function 


'// 状态码-对应意义
Function smpState(xCode,xName,xVal)
 Dim ac,an,i
 ac = split(xCode,";")
 an = split(xName,";")
 For i = 0 to uBound(ac)
   If cStr(ac(i))=cStr(xVal) Then
     smpState = an(i)
	 Exit Function
   End If 
 Next
 smpState = "(未知错误!)"
End Function
'// 调试状态：smcDebug = "Debug"
Function smpDebug(rArr,rText) 
  If smcDebug = "isDebug" Then
	Response.Write "<br>"&rArr(0)&"<br>"&rArr(1)&"<br>"&rArr(2)
	rText = Replace(Replace(rText,smpSoapHead,""),smpSoapFoot,"") 
	rText = Replace(Replace(rText,"http://www.w3.org/2001/",".../"),"http://schemas.xmlsoap.org/soap/",".../")
	Response.Write "<hr>"&smpXmlShow(rText)
  End If
End Function


'// 简易加密
Function smpEncode(xStr)   
Dim s,c1,c2,i
   For i=1 To Len(xStr)
     c1 = Mid(xStr,i,1)
	 c2 = Mid(smcECfg,(i Mod 8)+1,1)
	 s = s&cStr(Hex(Asc(c1)+Asc(c2)))
   Next
smpEncode = s
End Function
'// 简易解密
Function smpUncode(xStr)   
Dim s,j,c,n,i
   For i=1 To Len(xStr) Step 2
     j = (i+1)/2 'i = 1,3,5
	 c = Mid(smcECfg,(j Mod 8)+1,1)
	 n = CInt("&h"&Mid(xStr,i,2)) - Asc(c)
	 s = s&Chr(n)
   Next
smpUncode = s
End Function
'Response.Write "<br>"&smpEncode("peace")
'Response.Write "<br>"&smpUncode("A392A091A7A5AB96616363")


'// Show xml 代码
Function smpXmlShow(xText)
  Dim rText : rText = xText
  'rText = Replace(Replace(rText,smpSoapHead,""),smpSoapFoot,"") 
  rText = Replace(Replace(rText,"http://www.w3.org/2001/",".../"),"http://schemas.xmlsoap.org/soap/",".../")
  rText = Replace(rText,"><",">[bR]"&vbcrlf&"<")
  rText = Replace(rText,"<","&lt;")
  rText = Replace(rText,">","&gt;")
  rText = Replace(rText,"[bR]","<br>")
  smpXmlShow = rText
End Function
'// 获取 xml Node值 如<cMount>183</cMount>得到183
Function smpXmlNode(xHttp,xNode)
  Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
    xmlDOC.load(xHttp.responseXML)
    smpXmlNode = xmlDOC.documentElement.selectNodes("//"&xNode)(0).text
  Set xmlDOC = nothing
End Function
'// 获取 xml Val值
Function smpXmlVal(s,f) '用 smpXmlNode(xHttp,xNode) 代替
  Dim f1,f2,s0,p1,p2 : s0 = s
  f1 = "<"&f&">"
  f2 = "</"&f&">"
  p1 = inStr(s,f1) : If p1>0 Then p1=p1+Len(f1)
  p2 = inStr(s,f2)
  If p1>0 And p2>p1 Then
    s0 = Mid(s0,p1,p2-p1)
  Else
    s0 = ""
  End If
  smpXmlVal = s0
End Function
'// 获取 Html文本 
Function smpXmlText(xStr)
  Dim rExp,tStr : tStr=xStr            ' 建立变量。
  Set rExp = New RegExp                ' 建立正则表达式。
    rExp.Global = True                 ' 设置为全局
    rExp.Pattern = "<.*?>"              ' 设置模式。
    rExp.IgnoreCase = true             ' 设置是否区分大小写。
    smpXmlText = rExp.Replace(tStr,"") ' 作替换。
  Set rExp = Nothing
End Function



%>


