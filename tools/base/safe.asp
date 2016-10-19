<!--#include file="../../upfile/sys/para/keysafe.asp"-->
<%

'include <a href="../../upfile/sys/para/keysafe.asp">keysafe.asp</a>'
'ParFIPStr = "61.170.150.187|"
'ParFUrlStr = "/|"
'ParFKeyStr = "and 1=2|"

'Config_sfIPAct   = "Trace"                '禁IP操作：Trace:跟踪,不影响浏览; Stop:禁止IP,不能正常浏览; Exit:不启用；
'Config_sfInjAct  = "Trap"                 '防注入操作：Stop:停止操作,显示空白; Warn:警告信息; Trap:反攻击; Exit:不启用；
Config_sfPath    = Config_Path                '"/u/demo/"             '一般设置好
Config_sfPLogs   = "upfile/#dbf#/"         '记录文件路径
Config_sfPDir    = "tools/base/shome.asp"  '跳转文件路径
Config_sfNWarn   = 8                       '攻击次数达到多少次 执行
Config_sfLDays   = 3                       '每多少 天保存一个 记录文件 1,3,10, // 不用了 要么改/M/D/H 

'/// 系统函数，自动根据参数调用
Function sf_Guard()
  If Config_sfIPAct="Trace" Or Config_sfIPAct="Stop" Then
    sf_IPStop()
  End If
  If Config_sfInjAct="Trap" Or Config_sfInjAct="Warn" Or Config_sfInjAct="Stop" Then
    sf_InjStop()
  End If
  'Response.Write Config_sfInjAct&"<hr>------------------"
End Function

'/// 防注入函数
Function sf_InjStop()
  Dim uIP,vIP,pNow,nInj,isInj
  If Len(ParFKeyStr)<5 Then 
    Exit Function 'or <|
  End If
  isInj = sf_isInj()
  If isInj<>"" Then
	'Response.Write "<br>C:"&isInj '// for Test
	uIP = sf_IPReal()
	vIP = Left(uIP,inStrRev(uIP,".")-1)
	pNow = Request.ServerVariables("URL")&"?"&Request.ServerVariables("QUERY_STRING")
	Call sf_FLog("Inj",uIP,"[Check:"&isInj&"]",pNow)
	nInj = sf_noInj(vIP)
	If Config_sfInjAct="Warn" Then     '////// 警告
	  If Int(nInj)>=Int(Config_sfNWarn) Then
	    Call sf_FLog("Inj",uIP,"Warn","")
		Response.Redirect(Config_sfPath&Config_sfPDir&"?Act=InjWarn")
	  End If
	ElseIf Config_sfInjAct="Trap" Then '////// 陷阱  
	  If Int(nInj)=Config_sfNWarn Then
	    Call sf_FLog("Inj",uIP,"Warn","")
		Response.Redirect(Config_sfPath&Config_sfPDir&"?Act=InjWarn")
	  ElseIf Int(nInj)>=Int(Config_sfNWarn+3) Then
	    Call sf_FLog("Inj",uIP,"Trap","")
		Response.Redirect(Config_sfPath&Config_sfPDir&"?Act=InjTrap&uIP="&uIP&"&vIP="&vIP&"")
	  End If 
	End If                            '////// Stop停止不处理
	Response.End() 
  End If
End Function 

'/// 跟踪IP,或禁止IP函数
Function sf_IPStop()
  Dim uIP,vIP
  If Len(ParFIPStr)<5 Then 
    Exit Function '1.1.1
  End If
  uIP = sf_IPReal()
  vIP = Left(uIP,inStrRev(uIP,".")-1)
  If inStr(ParFIPStr,uIP)>0 Or inStr(ParFIPStr,vIP)>0 Then
    If Config_sfIPAct="Trace" Then
	  Call sf_FLog("IP.",uIP,"Trace","")
	  '/// 跟踪:记录后继续执行
	ElseIf Config_sfIPAct = "Stop" Then
	  Call sf_FLog("IP.",uIP,"Stop","")
	  Response.Redirect(Config_sfPath&Config_sfPDir&"?Act=IPStop")
	  Response.End() 
	End If 
  End If
End Function

'/// 检测 攻击次数
Function sf_noInj(xIP)
  Dim t,n
  Application("IPInj")=Application("IPInj")&xIP&"|"
  t = Application("IPInj")
  n = Len(t)-Len(Replace(t,xIP&"|",""))
  n = n/Len(xIP&"|")
  If Len(t)>1500 Then
    Application("IPInj") = Right(Application("IPInj"),800)
  End If
  sf_noInj = n
End Function
'Call sf_noInj(xIP)

'///  检测 是否攻击
' ParFKeyStr = ParFKeyStr&"and 1=2|and 1=1
Function sf_isInj()
  Dim gData,iItem,iData,a1,f,ai,b1,bn,bc,bi
    gData = "" 'Request.ServerVariables("QUERY_STRING")   '// url数据
  For Each iItem IN Request.QueryString()                 '// url数据
    iData = Request.QueryString(iItem)
    gData = gData&"|"&iData
  Next  
  'Response.Write "<br>A:"&gData '// for Test
  If gData="" Then
	For Each iItem IN Request.Form                        '// form数据
        iData = Request.Form(iItem)
	  If Len(iData)>2000 Then
        iData = Left(iData,800)&Right(iData,800)
		gData = gData&"|"&iData
	  End If 
    Next  
  End If  
  'Response.Write "<br>B:"&gData '// for Test
  a1 = Split(ParFKeyStr,"|")
  f = "" '记录注入特征字符
  For ai = 0 To uBound(a1)-1
    b1 = Split(a1(ai)," ")
	bn = uBound(b1)+1
	bc = 0 '
    For bi = 0 To uBound(b1)
      If inStr(lCase(gData),b1(bi))>0 Then
        bc = bc+1
      End If
    Next
	If bc = bn Then
	  f = a1(ai)
	  Exit For
	End If
  Next
  sf_isInj = f
End Function


' /// 格式： 2009-04-01 07:06:42 219.157.96.18 User-Member Act /news/news.asp?0ba69372cdb6fj680sfwn24d and(char(94)+user+char(94))>0 
' IP0904A.txt
Function sf_FLog(xFile,xUIP,xAct,xPage)
  Dim fn,fp,np,rp,dat
  fn = xFile&Get_FmtID("yymd","")&".txt"
  fp = Config_sfPath&Config_sfPLogs
  np = xPage
  If np="" Then
  np = Request.ServerVariables("URL")&"?"&Request.ServerVariables("QUERY_STRING")
  End If
  np = Replace(np,"%20"," ")
  rp = Request.Servervariables("HTTP_REFERER")
  dat = Now()&" "&xUIP&" "&xAct&" "&np&" [From] "&rp
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  fExist = fso.FileExists(Server.MapPath(fp&fn))
  If fExist Then
	Set fil2 = fso.OpenTextFile(Server.MapPath(fp&fn),8,True)
	fil2.WriteLine(dat) 
  Else
    Set fil2 = fso.CreateTextFile(Server.MapPath(fp&fn),True)
	fil2.Write(dat&vbcrlf)
  End If
  fil2.Close
  Set fso=Nothing
End Function
'Call sf_FLog("Inj",sf_IPReal(),"Trace","")


Function sf_IPReal() 
  Dim uIP
  uIP = Request.ServerVariables("REMOTE_ADDR")
  If uIP="" Then 
    uIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  End If
  sf_IPReal = uIP
End Function

%>
