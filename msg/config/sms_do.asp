<%

'说明: 此文件函数，为配合前台显示操作等相关函数，
'变量/函数 以do,log,get,chk,out等开头

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
Function outDecode(xStr)
  Dim i,a,s
  a = Split(xStr,";")
  For i=0 to uBound(a)
    If a(i)<>"" Then
	  ch = ChrW(a(i))
	  s=s&ch
	End If
  Next
  outDecode = s
End Function

' 0:True/False; 1:Count; 2:String
Function outTels() 
  smcOTMax = 20
  Dim r(2) :r(0)=false :r(1)=0 :r(2)=""
  Dim s,a,i,j
  s = smpFmtTNum(Request("tel")) '格式化
  a = Split(s,";") :s="" :j=0
  For i=0 To uBound(a)
    If smpChkTNum(a(i)) Then
	  If inStr(s,a(i))<=0 Then
	    j = j + 1
		s = s&a(i)&smcTSplit
		If j>=100 Then Exit For
	  End If
    End If
  Next
  If j>0 Then
    r(0) = true
	r(1) = j
	r(2) = Left(s,Len(s)-1)
  End If
  outTels = r
End Function

' 0:True/False; 1:Count; 2:String
Function outConts() 
  Dim r(2) :r(0)=false :r(1)=0 :r(2)=""
  Dim s,a,i,j,msg,m
    msg = Request("ct1")
	cs = Request("cs")
	If cs="" Or uCase(cs)="UTF-8" Then
	  If Len(msg)>70 Then
	    msg = Left(msg,70)
	  End If
	Else
	  'AscW(ch)ChrW(cu)
	  a = Split(msg,";") 
	  msg = "" :j = 0
	  For i=0 To uBound(a)
	    If isNumeric(a(i)) Then
		  msg = msg&ChrW(a(i))
	      j = j + 1
		  If j>=70 Then Exit For
		Else
		  Exit For
		End If
	  Next
	End If
	If Len(msg)>0 Then
	  r(0) = true
	  r(1) = 1
	  r(2) = msg
	End If
  outConts = r
End Function

'out发送短信 
Function outSend(xUser,xBalance,xType)
  Dim tInfo,tCont,r3(2),nTels,nCont,nFee,mStr
  tInfo = outTels()
  cInfo = outConts()
  If tInfo(0) And cInfo(0) Then
    nTels = tInfo(1) :tel = tInfo(2) 
	nCont = cInfo(1) :ct1 = cInfo(2)
	nFee = nTels*nCont
	If Int(xBalance)<Int(nFee) Then
	  r3(0) = "Error"
	  r3(1) = -41
	  r3(2) = "Error Balance!"
	ElseIf nTels>smcTMax1 Then
	  r3(0) = "Error"
	  r3(1) = -43
	  r3(2) = "Tels too Long!"
	Else ' Send(Out) out发送短信
	  res = smsSend(tInfo(2),cInfo(2),"","") 'sTime,sPort
	  If res(0) Then 
		Call doBalance("Member",xUser,nFee) '扣费
		r3(0) = "OK"
		r3(1) = nFee
		r3(2) = "Send OK!"
	  Else
		r3(0) = "Error"
		r3(1) = -46
		r3(2) = "Send Error!"
	  End If
	  Dim a,b,c
	  'a = Left(tInfo(2),3) :b = Right(tInfo(2),8)
	  'c = cStr(a+100)&cStr(b-11)
	  'Call logSend(c,"Test(OutAPI)自动",mStr,nFee,xUser)
	  mStr = res(0)&"(Out"&xType&");"&nTels&","&nCont&";Mem"&xType&""
	  Call logSend(tInfo(2),cInfo(2),mStr,nFee,xUser)
	End If
  Else
	  r3(0) = "Error"
	  r3(1) = -42
	  r3(2) = "Error Tel Number Or Message Content!"
  End If
  outSend = r3
End Function

'out修改SN 
Function outEdit(xUser,xSN)
  Dim r3(2),sql
  sql = "UPDATE SmsMember SET MemCode='"&xSN&"' WHERE MemID='"&xUser&"' "
  Call rs_DoSql(conn,sql)
  r3(0) = "OK"
  r3(1) = 0 '-51
  r3(2) = "Edit SN!"
  ' /// add Log?
  outEdit = r3
End Function

'outGetUrl
Function outGetUrl(xType)
  Dim pThis,pRefe,t
  pThis = Request.ServerVariables("HTTP_HOST")';localhost:240
  pThis = pThis&Request.ServerVariables("SCRIPT_NAME")'; /u/demo/tools/scan/aspcheck.asp 
  pRefe = Request.Servervariables("HTTP_REFERER")
  If inStr(pRefe,"?")>0 Then pRefe = Left(pRefe,inStr(pRefe,"?")-1)
  pRefe = Mid(pRefe,inStr(pRefe,"://")+3) 'http://,https://
  If inStr(xType,"/")>0 Then
    t = xType
  ElseIf xType="this" Then
    t = pThis
  ElseIf xType="chk" Then
   If pThis=pRefe Then
    t = Request(Request("sn"))
   Else
    t = pRefe
   End If
  Else 'xType="" Or xType="rp"
    t = pRefe
  End If 
  'Response.Write h
  outGetUrl = t
End Function
'out加密sn
Function outEncSN(xid,xtm,xsn,xUrl)
  Dim t,ytm
  ytm = outFmtTime(xtm)
  t = xid&"+"&xsn&"+"&ytm&"+"&outGetUrl(xUrl)
  'Response.Write h
  outEncSN = MD5_32(t)
End Function
'out编码中文
Function outEncMsg(xStr,xSet,xNow)
  Session.Codepage = xSet
  outEncMsg = Server.URLEncode(xStr)
  Session.Codepage = xNow'65001'还原当前会话编码
End Function
' 检查加密串 id+sn+tm+url 
Function outChkCode(xsn,xid,xtm,xCode,xUrl,xFlag) 
  Dim rp,r,t,ytm : r = 0
  ytm = outFmtTime(xtm)
	If xFlag="N" Then 'Stop
      r = -32
	Else
	  rp = outGetUrl("chk") 'Request("rp")
	  'Response.Write "<br>"&xUrl&"<br>"&rp
	  If rp<>xUrl Or xUrl="" Then
        r = -33
		'echo rp&vbcrlf&xUrl
	  End If
	  t = xid&"+"&xCode&"+"&ytm&"+"&xUrl
	  'echo "<br>:"&t
	  'echo "<br>:"&xsn
	  'echo "<br>:"&MD5_32(t)
	  If xsn<>MD5_32(t) Then
        r = -34
	  End If
	End If 
  outChkCode = r
End Function
' 检查超时时间
Function outChkTime(xAct,xTime) 
  Dim n,f : f = true
  n = Abs(DateDiff("n",Now,xTime))
  If n>1440 Then '超时一天，Error 一天:24x60=1440(min)
    f = false
  ElseIf inStr("(Balance,Readme,)",xAct)>0 And n>180 Then '超时180分钟
    f = false
  ElseIf inStr("(SendMsg,SendOut,SendODo,EditSN,)",xAct)>0 And n>90 Then '超时90分钟
    f = false
  ElseIf n>120 Then '超时120分钟
    f = false
  End If
  outChkTime = f
End Function

'发送短信 每次发送Max1个号码
Function doSInn(xCount,xTels,xCont,xTime,xPort)
  Dim msg,nMax,aTel,i,j,s :msg=""
  If xCount<=smcTMax1 Then
    msg = smsSend(xTels,xCont,xTime,xPort)
	'Call echo("1."&xTels)
  Else
    nMax = Int(smcTMax1*0.8)
	aTel = Split(xTels,smcTSplit)
	j = 0 :s = ""
	For i=0 To uBound(aTel)
	  j = j+1
	  If j=nMax Or j=uBound(aTel) Then
		s = s&aTel(i)&""
		msg = smsSend(s,xCont,xTime,xPort) 'msg&"<br>"&
		'Call echo("2."&s)
		j = 0 :s = ""
	  Else
		s = s&aTel(i)&","
	  End If
	Next
  End If
  doSInn = msg
End Function
'短信群发 每次发送Max1个号码
Function doGroup() 
  Dim tNumb,msg,nMax,aTel,j,s,i
  tNumb = getTels("") :msg=""
  If NOT tNumb(0) Then
	msg = "发送失败;<br>0个号码,扣费:0条"
  ElseIf tNumb(1)<=smcTMax1 Then
	msg = doSend(sndType,sndUser,tNumb(2))
  Else
	nMax = Int(smcTMax1*0.8)
	aTel = Split(tNumb(2),smcTSplit)
	j = 0 :s = ""
	For i=0 To uBound(aTel)
	  j = j+1
	  If j=nMax Or j=uBound(aTel) Then
		s = s&aTel(i)&"" ':echo "2."&s 
		msg = msg&"<br>"&doSend(sndType,sndUser,s)
		j = 0 :s = ""
	  Else
		s = s&aTel(i)&";"
	  End If
	Next
  End If
  doGroup = msg
End Function
'发送短信 每次发送Max2个号码
Function doSend(xType,xUser,xTels) 
  Call chkReload()
  Dim nBalance,tInfo,cInfo,cMode,emsg,eflg,nTels,nCont,nFee
  Dim sTime,sPort,sUser,res,mStr,mFee,sFlg,i,t,aCont
  nBalance = doBalance(xType,xUser,-1) '获取用户余额
  tInfo = getTels(xTels) '"tNumb"
  cInfo = getConts("tCont","Auto")
  cMode = Request("tMode")
  emsg = "" : eflg=false 
  If tInfo(0) And cInfo(0) Then
    nTels = tInfo(1)
	nCont = cInfo(1)
	nFee = nTels*nCont ':Response.Write "<br>22"&nBalance &":"& nFee ':Response.End()
	If Int(nBalance)<Int(nFee) Then
	  emsg = "发送失败,余额不足"
	  eflg = true
	ElseIf nTels>smcTMax2 Then
	  emsg = "发送失败,号码太多"
	  eflg = true
	  'echo nTels
	Else
	    'xxx = Send Start
		sTime = Request("tTime")
		sPort = Request("tPort")
		If xUser="" Then
		  If xType="Admin" Then
			sUser = Session("UsrID")
		  Else
			sUser = Session("MemID")
		  End If
		Else
		  sUser = xUser
		End If
		'发送,扣费,写Log
		If cMode="Long" Then '长信息
		  If smcCLong Then '支持长内容
			'////////////////////// Send(Long) 长信息 /////////////////////////////////////
			res = doSInn(tInfo(1),tInfo(2),cInfo(2),sTime,sPort) 
			If res(0) Then 
			  Call doBalance(xType,xUser,nFee) '扣费
			Else
			  nFee = 0 
			End If
			mStr = res(0)&"(Long);"&nTels&","&nCont&";"&xType
			Call logSend(tInfo(2),cInfo(2),mStr,nFee,xUser)
			doSend = res(2)&"(长信息发送);<br>"&nTels&"个号码,"&nCont&"条信息,扣费:"&nFee&"条;<br>发送跟踪标记:"&mStr&""
			'////////////////////// Send(Long) End /////////////////////////////////////  
		  Else
			emsg = "发送失败, 不支持长内容"
			eflg = true
		  End If
		ElseIf cMode="More" Then '多信息
		  If smcCMul Then '支持多条内容
			'////////////////////// Send(Batch) 多条内容 /////////////////////////////////////
			res = doSInn(tInfo(1),tInfo(2),cInfo(2),sTime,sPort) 
			If res(0) Then 
			  Call doBalance(xType,xUser,nFee) '扣费
			Else
			  nFee = 0 
			End If
			mStr = res(0)&"(Batch);"&nTels&","&nCont&";"&xType
			Call logSend(tInfo(2),cInfo(2),mStr,nFee,xUser)
			doSend = res(2)&"(批量发送);<br>"&nTels&"个号码,"&nCont&"条信息,扣费:"&nFee&"条;<br>发送跟踪标记:"&mStr&""
			'////////////////////// Send(Batch) End ///////////////////////////////////// 
		  Else
			'////////////////////// Send(BFor) 多条内容 /////////////////////////////////////
			aCont = Split(cInfo(2),smcCSplit) 
			nFee = 0 :mFee = 0 :sFlg = "" '成功/失败
			For i=0 To uBound(aCont)
			  res = doSInn(tInfo(1),tInfo(2),aCont(i),sTime,sPort) 
			  If res(0) Then 
				Call doBalance(xType,xUser,nFee) '扣费
				nFee = nFee + nTels
				sFlg = sFlg&"T"
			  Else
				mFee = mFee + nTels
				sFlg = sFlg&"F"
			  End If
			Next
			mStr = sFlg&"(BFor);"&nTels&","&nCont&"("&mFee&");"&xType
			If mFee=0 Then 
			  t = "成功！"
			ElseIf nFee=0 Then 
			  t = "失败！"
			Else
			  t = "部分信息发送成功！"
			End If
			Call logSend(tInfo(2),cInfo(2),mStr,nFee,xUser)
			doSend = t&"(循环发送);<br>"&nTels&"个号码,"&nCont&"条信息,扣费:"&nFee&"条(失败:"&mFee&"条);<br>发送跟踪标记:"&mStr&""
			'////////////////////// Send(BFor) End ///////////////////////////////////// 
		  End If
		Else 'cMode="One" '一条信息
		  '////////////////////// Send(Single) 单条内容 /////////////////////////////////////
		  res = doSInn(tInfo(1),tInfo(2),cInfo(2),sTime,sPort)  
		  If res(0) Then 
			Call doBalance(xType,xUser,nFee) '扣费
		  Else
			nFee = 0 
		  End If
		  mStr = res(0)&"(One);"&nTels&","&nCont&";"&xType
		  Call logSend(tInfo(2),cInfo(2),mStr,nFee,xUser)
		  doSend = res(2)&"(单条信息发送);<br>"&nTels&"个号码,扣费:"&nFee&"条;<br>发送跟踪标记:"&mStr&""
		  '////////////////////// Send(Single) End /////////////////////////////////////
		End If
	    'xxx = Send End
	End If
  Else
	emsg = "发送失败, 空内容或空号码"
	eflg = true
  End If
  If eflg Then
	Response.Redirect "../page/message.asp?mType=Error&msg="&emsg&"&goPage=goRef"
	Response.End()   
  End If  
End Function

'发送测试
Function doTest()
  Call chkReload()
  Dim tNumb,tCont,res,nCnt,mStr
  aNumb = getTels("tNumb") :tNumb = aNumb(2)
  tCont = RequestS("tCont",3,2400) :nCnt = aNumb(1)*smpCntCont(tCont)
  res = smsTest(tNumb,tCont)
  mStr = res(0)&"(Test);("&aNumb(1)&","&nCnt&");"&smcDMod
  Call logSend(tNumb,tCont,mStr,nCnt,Session("UsrID"))
  doTest = res
End Function

'获取余额/余额扣费 
'xType (SmsAPI),Admin(不限制),Inner,Message,Member(SmsMember)
'xCount -1:获取余额 >0:余额扣费 
Function doBalance(xType,xUser,xCount)
  Dim uTab,uCell,uKey,sUser,res,r,sql
  '设置数据表
  If xType="Member" Then
	uTab = "SmsMember"
	uCell = "MemBalance"
	uKey = "MemID"
  Else
	uTab = "AdmUser"&Adm_aUser&""
	uCell = "UsrBalance"
	uKey = "UsrID"
  End If
  '检查用户
  If xUser="" Then
	If xType="Admin" Then
	  sUser = Session("UsrID")
	ElseIf xType="Inner" Then
	  sUser = Session("InnID")
	ElseIf xType="Message" Then
	  sUser = Session("SmsID")
	ElseIf xType="Member" Then
	  sUser = Session("MemID")
	End If
  Else
	sUser = xUser
  End If
  'Do 
  If xType="(SmsAPI)" And Int(xCount)=-1 Then '获取System余额
    res = smsBalance()
    doBalance = res(1) 
  ElseIf Int(xCount)=-1 Then '获取User余额
	r = rs_Val(conn,"SELECT "&uCell&" FROM "&uTab&" WHERE "&uKey&"='"&xUser&"' ")
	doBalance = RequestSafe(r,"N",0)
  ElseIf xCount>0 Then 'User余额扣费
	sql = "UPDATE "&uTab&" SET "&uCell&" = "&uCell&"-" &xCount&" WHERE "&uKey&"='"&xUser&"' "
	Call rs_DoSql(conn,sql)
  End If
End Function

' 0:True/False; 1:Count; 2:String
Function getTGrp(xPara) '"tNumb"
  Dim r(2) :r(0)=false :r(1)=0 :r(2)=""
  Dim s,a,i,j
  s = Replace(Replace(RequestS(xPara,"",120123)," ",""),"	","")' 得到数据, 去掉空格
  s = Replace(Replace(s,vbcr,";"),vblf,";") ' 回车换行
  a = Split(s,";") :s="" :j=0
  For i=0 To uBound(a)
    If smpChkTGrp(a(i)) Then
	  If inStr(s,a(i))<=0 Then
	    j = j + 1
		s = s&a(i)&vbcrlf
	  End If
    End If
  Next
  If j>0 Then
    r(0) = true
	r(1) = j
	r(2) = Left(s,Len(s)-2)
  End If
  getTGrp = r
End Function

' 0:True/False; 1:Count; 2:String
Function getTels(xPara) '"tNumb"
  Dim r(2) :r(0)=false :r(1)=0 :r(2)=""
  Dim s,a,i,j
  If xPara="" Or xPara="tNumb" Then
	s = smpFmtTNum(Request("tNumb")) 
  ElseIf Len(xPara)<8 Then 
    s = smpFmtTNum(Request(xPara)) 
  Else
	s = smpFmtTNum(xPara) '格式化 
  End If
  a = Split(s,";") :s="" :j=0
  For i=0 To uBound(a)
    If smpChkTNum(a(i)) Then
	  If inStr(s,a(i))<=0 Then
	    j = j + 1
		s = s&a(i)&smcTSplit
	  End If
    End If
  Next
  If j>0 Then
    r(0) = true
	r(1) = j
	r(2) = Left(s,Len(s)-1)
  End If
  getTels = r
End Function

' 0:True/False; 1:Count; 2:String
Function getConts(xPara,xMode) 
  Dim r(2) :r(0)=false :r(1)=0 :r(2)=""
  Dim s,a,i,j,iMsg,m
  If xMode="Auto" Then
    m = Request("tMode")
  Else
    m = xMode
  End If
  If m="Long" Then '长信息
	iMsg = Request(xPara&"1")
	If Len(iMsg)>255 Then
	  iMsg = Left(iMsg,70)
	End If
	If Len(iMsg)>0 Then
	  r(0) = true
	  r(1) = smpCntCont(iMsg)
	  r(2) = iMsg
	End If
  ElseIf m="More" Then '多信息
	s = "" :j=0
	For i=1 To 9
	  iMsg = Request(xPara&i) 
	  If Len(iMsg)>70 Then
		iMsg = Left(iMsg,70)
	  End If
	  If iMsg<>"" Then
		j = j + 1
		s = s&iMsg&smcCSplit
	  End If
	Next
	If j>0 Then
	  r(0) = true
	  r(1) = j
	  r(2) = Left(s,Len(s)-Len(smcCSplit))
	End If
  Else 'm="One" '一条信息
    iMsg = Request(xPara&"1")
	If Len(iMsg)>70 Then
	  iMsg = Left(iMsg,70)
	End If
	If Len(iMsg)>0 Then
	  r(0) = true
	  r(1) = 1
	  r(2) = iMsg
	End If
  End If
  getConts = r
End Function

'发送记录
Function logSend(xNumb,xCont,xRes,xRCnt,xUser)
  Dim sql '255
  If Len(xCont)>255 Then xCont=Left(xCont,150)&" ... ... "&Right(xCont,80)&"("&Len(xCont)&")"
  If Len(xNumb)>255 Then xNumb=Left(xNumb,150)&" ... ... "&Right(xNumb,80)&"("&Len(xNumb)&")"
  xCont = RequestSafe(Replace(xCont,smcCSplit,"(msg)"),"C",510) 
  xNumb = RequestSafe(xNumb,"C",510)
  sql =      "INSERT INTO [SmsSend] ("
  sql = sql& "  LogID, LogTels, LogCont" 
  sql = sql& ", LogRes, LogRCnt"
  sql = sql& ", LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES("
  sql = sql& "  '" & Get_AutoID(24) &"', '" & xNumb &"', '" & xCont &"'"
  sql = sql& ", '"& xRes &"', " & xRCnt &""
  sql = sql& ",'"&Get_CIP()&"','"&xUser&"','"&Now()&"'"
  sql = sql& ")" ':Response.Write sql
  Call rs_DoSql(conn,sql)
End Function

'充值记录
Function logCharge(xUser,xMoney,xCount,xNote)
  Dim sql
  sql =      "INSERT INTO [SmsCharge] ("
  sql = sql& "  LogID, LogMoney, LogCount" 
  sql = sql& ", LogUser, LogNote"
  sql = sql& ", LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES("
  sql = sql& "  '" & Get_AutoID(24) &"', " & xMoney &", " & xCount &""
  sql = sql& ", '"&xUser&"', '" & RequestSafe(xNote,"C",510) &"'"
  sql = sql& ",'"&Get_CIP()&"','"&Session("UsrID")&"~"&Session("MemID")&"','"&Now()&"'"
  sql = sql& ")" ':Response.Write sql
  Call rs_DoSql(conn,sql)
End Function
  
'重复提交检查
Function chkReload()
  Dim cCode
  cCode = uCase(Request("ChkCode"))
  If cCode="" Or cCode<>Session("ChkCode") Then
	Response.Redirect "../page/message.asp?mType=Error&msg=请不要重复提交&goPage=goRef"
	Response.End()   
  End If
End Function

'不支持函数检查
Function chkDTabs(xAct)
  If inStr(smcDTabs,xAct)>0 Then
    Response.Write "disabled='disabled'"
	msg = "此接口 不支持此功能！"
  End If
End Function

%>


