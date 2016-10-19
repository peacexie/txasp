<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->
<!--#include file="../../upfile/sys/pcfg/sms.asp"-->
<%

Call Chk_Perm1("SysTools","")  '请设置好访问权限

'smaTitle = "短信接口API"     '短信接口
'smaState = "isOpen"          '状态(isStop,isOpen)
'smaUser = "(id)"             '接口帐号(预设)
'smaCode = "(sn)"             '接口SN(预设)
'smaAdda = "(ad1)"            '接口地址(远程)
'smaAddr = "(ad2)"            '发送地址(本地)
'smaTels = "(tel)"            '发送号码(网管)

Dim id,snOrg
id = smaUser 
snOrg = smaCode 
smsABase = Left(smaAdda,inStr(smaAdda,"/msg/user/"))

'时间格式化
Function outFmtTime(xTime)
  Dim dy,dm,dd,th,tn,ts
  If Len(xTime)=19 Then
   outFmtTime = xTime
  Else
   dy = DatePart("yyyy",xTime)
   dm = DatePart("m",xTime) :If Len(dm)=1 Then dm="0"&CStr(dm)
   dd = DatePart("d",xTime) :If Len(dd)=1 Then dd="0"&CStr(dd)
   th = DatePart("h",xTime) :If Len(th)=1 Then th="0"&CStr(th)
   tn = DatePart("n",xTime) :If Len(tn)=1 Then tn="0"&CStr(tn)
   ts = DatePart("s",xTime) :If Len(ts)=1 Then ts="0"&CStr(ts)
   outFmtTime = ""&dy&"-"&dm&"-"&dd&" "&th&":"&tn&":"&ts&""
  End If
End Function
'echo outFmtTime("2011-4-25 4:7:2")
'内容编码,解码
Function outEncode(xStr)
   Dim i,ch,cx,s
   For i=1 to len(xStr)
	  ch=Mid(xStr,i,1)
	  cx=AscW(ch) 'Hex()
	  if cx<0 Then cx=cx+65536
	  s=s&cx&";"
   Next
   outEncode = s
End Function
'sTest = "测试byPeace鸽子!"
'echo outEncode(sTest)
Function outDecode(xStr)
  Dim i,a,ch,s
  a = Split(xStr,";")
  For i=0 to uBound(a)
    If a(i)<>"" Then
	  ch = ChrW(a(i))
	  s=s&ch
	End If
  Next
  outDecode = s
End Function

Function outEncSN(xid,xtm,xsn)
  Dim t,ytm
  ytm = outFmtTime(xtm)
  t = xid&"+"&xsn&"+"&ytm&"+"&smaAddr
  'Response.Write h
  outEncSN = MD5_32(t)
End Function

%>