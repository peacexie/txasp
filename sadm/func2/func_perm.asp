<%

Dim ckv2_Inc,ckv3_srv,ckv5_out,ckv4_code,ckv1_enc,ckv0_net
ckv2_Inc = "<!--\s*#include\s*(virtual|file)\s*=\s*"".*"""
ckv3_srv = "<"&"%(\s|\S)*?Server\.(Execute|Transfer|CreateObject)(\s|\S)*?%"&">"
ckv5_out = "<script[^>]*?>(\s|\S)*?<\/script[^>]*?>|"&_
          "<iframe[^>]*?>(\s|\S)*?<\/iframe[^>]*?>|"&_
		  "<object[^>]*?>(\s|\S)*?<\/object[^>]*?>|"&_ 
		  "<applet[^>]*?>(\s|\S)*?<\/applet[^>]*?>|"&_
		  "<frameset[^>]*?>(\s|\S)*?<\/frameset[^>]*?>"
ckv4_code = "<"&"%(\s|\S)*?(OpenTextFile\(|CreateTextFile\(|eval |execute |eval\(|execute\()(\s|\S)*?%"&">"
ckv1_enc = "<"&"%(\s|\S)*?(VBScript|JScript)\.Enc"&"ode(\s|\S)*?\#\@\~\^(\s|\S)*?\=\=\^\#\~\@(\s|\S)*?%"&">"
ckv0_net = "<"&"%@\s*Page[^>]*?%"&">|<"&"%@\s*import[^>]*?%"&">"


'// 显示额外配置入口
'// 相当与设置Config_Mode=isExpert 
Function Chk_PermSP()
  spTab = "(peace,yufish,xieys,)"
  if inStr(spTab,Session("UsrID")&",")>0 Then
    Chk_PermSP = true
  Else
    Chk_PermSP = false
  End If
End Function

'// 有审核(Check) 列表
Function Flg_PCheck()
  'Response.Write xDep&xMod
  Dim sPerm,r
  sPerm=Session(UsrPStr)
  If inStr(sPerm,"{Admin}") > 0 Then
    r = true
  ElseIf inStr(sPerm,"(SetShow)") > 0 Then
    r = true
  Else
    r = false
  End If
 Flg_PCheck = r
End Function

'// 有权限(List) 列表
Function Flg_PList()
  'Response.Write xDep&xMod
  Dim sPerm,r
  sPerm=Session(UsrPStr)
  If inStr(sPerm,"{Admin}") > 0 Then
    r = true
  ElseIf inStr(sPerm,"(SetOher)") > 0 Then
    r = true
  Else
    r = false
  End If
 Flg_PList = r
End Function

'// 有权限(部门) 删除/修改
Function Flg_PDep(xDep,xMod)
  'Response.Write xDep&xMod
  Dim r,dFlag : dFlag = Eval(xMod&"Typ2Code")
  If dFlag<>"(Depart)" Then
    r = true
  Else
   If inStr(Session(UsrPStr),"{Admin}")>0 Then
    r = true
   ElseIf inStr(Session(UsrPStr),"("&xDep&")")>0 Then
    r = true
   Else
    r = false
   End If
 End If
 Flg_PDep = r
End Function

'// 有权限 删除/修改
Function Flg_PUser(f,xUser)
  Dim r 'If f="(Inn)" OR f="(Mem)" Then
  If inStr(Session(UsrPStr),"{Admin}")>0 Then
    r = true
  ElseIf inStr(Session(UsrPStr),"(SetOher)")>0 Then
    r = true '了不要//根本不类别出来
  ElseIf xUser=Session("UsrID") Or xUser=Session("InnID") Or xUser=Session("MemID") Then
    r = true
  Else
    r = false
  End If
  Flg_PUser = r
End Function

'//  LogAUser = Get_PUser(PrmFlag)
Function Get_PUser(f) 
  Dim u : u = "(Nul)"
  If f = "(Inn)" Then
    u = Session("InnID")
  ElseIf f = "(Mem)" Then
    u = Session("MemID")
  Else
    u = Session("UsrID") '"(Nul)"
  End If
  Get_PUser = u
End Function

' 指定检测，判断权限，优先级：Adm,Mem,Inn,3
Function Chk_Perm9(xSyst,xFlag) 
  Dim sPerm : sPerm="" 'Response.Write xSyst
  If xFlag="(Adm)" Then 
    Call Chk_Perm1(xSyst,xAct)
  ElseIf xFlag="(Mem)" Then 
    Call Chk_Perm2(xSyst,xAct)
  ElseIf xFlag="(Inn)" Then 
    Call Chk_Perm3(xSyst,xAct)
  ElseIf xFlag="3" Then
   sPerm=Session(UsrPStr)&""&Session("MemPerm")&""&Session("InnPerm")
   If inStr(UsrPerm,"{Admin}") > 0 Then
	 '// Continue for Admin
   ElseIf xSyst="" And LEN(Session(UsrPStr)&"")>3 Then
	 '// Continue for Null xSyst
   ElseIf inStr(sPerm,xSyst)>0 Then
	 '// Continue for xSyst(Mod)
   Else
	 Response.Write "(No Perm!)"
	 Response.End()
   End If
  Else
    Call Chk_Perm1(xSyst,xAct)
  End If
End Function

' /////////////////////////////////////////////////////// 


Function Chk_PSub(xPath) 
  ' Const Config_Path  = "/u/test21/" Session("Pub_Subs") = Config_Path
 If Len(Config_Path)>5 And lCase(Left(Config_Path,3))="/u/" Then
   'Chk_URL()
   Dim vPath : vPath = Request.ServerVariables("URL")
   Dim vSubs : vSubs = Session("Pub_Subs")&""
   If inStr(vPath,vSubs)<=0 OR Len(vSubs)<5 Then Response.Redirect xPath
 End If
End Function

Function Chk_Perm1(xSTID,xAct) 
  'exit function 
  Dim Obj,UsrPerm
  UsrPerm = Session(UsrPStr)&""
  Call Chk_PSub(Config_Path&"sadm/user/error.asp?send=NoLogin")
  If Session("UsrID")&"" = "" Then 
    Response.Redirect Config_Path&"sadm/user/error.asp?send=NoLogin"
  End If
    Obj = false
  If inStr(UsrPerm,"{Admin}") > 0 Then
    Obj = true
  Else
    Obj = Chk_Perm0(UsrPerm,xSTID,xAct) 
  End If
  If NOT Obj Then 
    Response.Redirect Config_Path&"sadm/user/error.asp?send=NoPerm&SysID="&xSTID&"&Act="&xAct&""
  End If
  Chk_Perm1 = Obj
End Function

Function Chk_Perm2(xSTID,xAct) 
  'exit function 
  Call Chk_PSub(Config_Path&"member/login.asp?send=NoLogin")
  If Session("MemID")&"" = "" Then 
    Dim goDir : goDir=""
    If xAct<>"" Then goDir = "&goPage="&xAct
    Response.Redirect Config_Path&"member/login.asp?send=NoLogin"&goDir
  End If
  If NOT Chk_Perm0(Session("MemPerm")&"",xSTID,xAct) Then 
    Response.Redirect Config_Path&"member/info/error.asp?send=NoPerm&SysID="&xSTID&"&Act="&xAct&""
  End If
End Function

Function Chk_Perm3(xSTID,xAct) 
  'exit function 
  Call Chk_PSub(Config_Path&"ext/login.asp?send=NoLogin")
  If Session("InnID")&"" = "" Then 
    Dim goDir : goDir=""
    If xAct<>"" Then goDir = "&goPage="&xAct
    Response.Redirect Config_Path&"ext/login.asp?send=NoLogin"&goDir
  End If
  If NOT Chk_Perm0(Session("InnPerm")&"",xSTID,xAct) Then 
    Response.Redirect Config_Path&"ext/login.asp?send=NoPerm&goPage=goRef"
  End If
End Function

Function Chk_Perm0(xPerm,xSTID,xAct) 
  Dim r,UsrPerm 
  UsrPerm = Session(UsrPStr)&""
  If inStr(UsrPerm,"{Admin}") > 0 Then
    r = true
  ElseIf xSTID="" AND LEN(xPerm)>3 Then
    r = true
  ElseIf xSTID<>"" AND inStr(xPerm,xSTID)>0 Then
    r = true
  Else
    r = false
  End If
  Chk_Perm0 = r
End Function


'' "dg.gd.cn/business/user/info_add.asp" ''特定字串
'' "" '本域名
Function Chk_URL3(xUrl)
Dim PagePrev,PageThis,nPos,mPos,e1Str,e2Str,rObj
  PagePrev = Request.Servervariables("HTTP_REFERER")  '' http://localhost:306/xbak/t01.asp?c=Test03; 
  PagePrev = Replace(PagePrev,"http://","")           '' 去http://,含www.含端口
  PageThis = Request.ServerVariables("HTTP_HOST")     '' 不含http:// localhost:240 www.dgchr.com
  'Response.Write "<br>URL3A:"&PagePrev&"<br>"&PageThis '测试
	If Left(PagePrev,Len(PageThis))<>PageThis Then
	  Response.Redirect "/"
	End If
	'' Config_URL, http://localhost:256/u/demo/
	nPos = inStr(xUrl,"/")         '' dg.gd.cn/business/user/info_add.asp
	e1Str = Mid(xUrl,nPos+1)       '' business/user/info_add.asp
	mPos = inStr(PagePrev,"/")     '' localhost:306/xbak/t01.asp?c=Test03;
	e2Str = Mid(PagePrev,mPos+1)   '' xbak/t01.asp?c=Test03
	rObj = "OK"
	'Response.Write "<br>URL3B:"&e1Str&"<br>"&e2Str '测试
	If Left(e2Str,Len(e1Str))<>e1Str Then
	  rObj = "eUrl" 'Response.Redirect "/"
	End If
Chk_URL3 = rObj
End Function

Function Chk_URL()
'exit function  
Dim PagePrev,PageThis
	PagePrev = Request.Servervariables("HTTP_REFERER") '含http://
	PageThis = Request.ServerVariables("HTTP_HOST") '不含http://
    PagePrev = Replace(PagePrev,"http://","")
	'Response.Write "<br>URL:"&PagePrev&"<br>"&PageThis '测试
	If Left(PagePrev,Len(PageThis))<>PageThis Then
	  'Call Add_Log(conn,Session("UsrID"),"[Chk_URL()]","[ALL]",PagePrev&"-"&PageThis)
	  Response.Redirect "/"
	End If
End Function

'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
'Call ChkSpider(1) 1直接截断 0 输出检测结果
'------------------------------------------------
Function ChkSpider(tf)
  Dim ArrSpider,Agent_User,i
  '要屏蔽的搜索引擎的标志可以自己在下面加.
  ArrSpider=Array("Googlebot","Baiduspider","Yahoo Slurp","Yahoo! Slurp China","Sina Iask Spider","Bloglines","jakarta","httpclient","soso","twiceler","GouGou","MSNBot")
  Agent_User=LCase(Request.ServerVariables("HTTP_USER_AGENT"))
  'Agent_User=LCase("Test Baiduspider Test")
  ChkSpider = False
  For i = 0 To UBound(ArrSpider)
	If InStr(Agent_User,LCase(ArrSpider(i))) Then
	  ChkSpider = True
	  If tf=1 Then
		Response.Clear()
		Response.Status = 404
		Response.Write "<h1>404 Not Found</h1>"
		Response.End
	  End If
	  Exit Function
	End If
  Next
End Function

'------------------------------------------------
'检测木马
'Call ChkTrojan(xStr,"") 
'发现用正则效率还慢很多...如下
'f1 = Check_RTest(str,ckv2_Inc)
'f2 = Check_RTest(str,ckv3_srv)
'f3 = Check_RTest(str,ckv5_out)
'f4 = Check_RTest(str,ckv4_code)
'f5 = Check_RTest(str,ckv1_enc)
'f6 = Check_RTest(str,ckv0_net)
'------------------------------------------------
Function ChkTrojan(xStr,xMod)
  Dim ckuf_asp1,ckuf_asp2,ckuf_asp,ckuf_flag
  Dim str,mods,cku5_out
  ckuf_asp1 = "<"&"%" : ckuf_asp2 = "%"&">"
  ckuf_asp = false    : ckuf_flag = ""
  str = lCase(xStr)   : mods = lCase(xMod)
  If mods="" Then mods="inc,srv,code,enc1,enc2,net,out"
  If inStr(str,"<"&"%")>0 And inStr(str,"%"&">")>0 Then
	ckuf_asp = true
  End If
  If inStr(mods,"inc")>0 Then 'Inc:<!--|include|(virtual|file)-->
	If inStr(str,"<!--")>0 And inStr(str,"-->")>0 And inStr(str,"include ")>0 And (inStr(str," virtual")>0 Or inStr(str," file")>0) Then
	  ckuf_flag = "inc"
	End If 
  End If 
  If inStr(mods,"srv")>0 Then 'srv:Server.(Execute|Transfer|CreateObject(
	If ckuf_asp And inStr(str,"server.")>0 And (inStr(str,"execute")>0 Or inStr(str,"transfer")>0 Or inStr(str,"createobject")>0) Then
	  ckuf_flag = "srv"
	End If 
  End If 
  If inStr(mods,"code")>0 Then 'code:OpenTextFile|CreateTextFile|eval |execute |eval(|execute(
	If ckuf_asp And (inStr(str,"opentextfile")>0 Or inStr(str,"createtextfile")>0 Or inStr(str,"eval ")>0 Or inStr(str,"execute ")>0 Or inStr(str,"eval(")>0 Or inStr(str,"execute(")>0) Then
	  ckuf_flag = "code"
	End If 
  End If 
  If inStr(mods,"enc1")>0 Then 'enc1:language = VBScript.Encode/JScript.Encode
	If ckuf_asp And inStr(str,"language")>0 And (inStr(str,"jscript.encode")>0 Or inStr(str,"vbscript.encode")>0) Then
	  ckuf_flag = "enc1"
	End If 
  End If 
  If inStr(mods,"enc2")>0 Then 'enc2:#@~^bgEAAA==...@#@&E3EAAA==^#~@
	If ckuf_asp And inStr(str,"#@~^")>0 And inStr(str,"==^#~@")>0 Then
	  ckuf_flag = "enc2"
	End If 
  End If 
  If inStr(mods,"net")>0 Then 'net:@Page|@import for .net
	If ckuf_asp And inStr(str," page ")>0 And inStr(str," import ")>0 Then
	  ckuf_flag = "net"
	End If 
  End If 
  If inStr(mods,"out")>0 Then 'out:script|iframe|object|applet|frameset
	cku5_out = Split("script|iframe|object|applet|frameset","|")
	For i=0 To uBound(cku5_out)
	  If inStr(str,"<"&cku5_out(i))>0 And inStr(str,"</"&cku5_out(i))>0 Then
		ckuf_flag = "out"
		Exit For
	  End If
	Next 
  End If 
  ChkTrojan = ckuf_flag
End Function


%>
