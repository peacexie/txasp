<%

'<!--/*  sms_cfg.asp 定义的配置
'亿美软通 接口 
smcMClass = "" 'EUCPComm.dll
smcUser = "3SDK-EGD-0130-LKROR" ' 用户名/序列号 0SDK-EMY-0130-LHZTO,3SDK-EGD-0130-LKROR
'smcUEID = "" ' 企业ID
smcPass = "187740" '密码 145505,187740
smcPase = "??A4026543695874F43244CDFB58149C" '
smcTSplit = "," ' 号码分割号
smcTMax1 = 1000'3,1000 ' 接口群发号码数
smcTMax2 = 3200'4,9600 ' 系统群发号码数
smcCLong = false ' true,false长内容
smcCMul = false ' true,false多条内容
smcDMod = "emay.cn亿美(Http)"
smcDTabs = "(Charge;Report;)" 'Charge;Receive;Report
smcSTime = true ' true,false定时发送
smcSPort = false ' true,false子端口
'*/-->

'说明: 此文件为 调用接口的函数 
'不同接口使用不同函数， smsSend()函数名不变，内容根据需要改变；
'所有 变量/函数 以sms开头

'接口公共数据
smsHttpUrl = "http://sdkhttp.eucp.b2m.cn/sdkproxy/" 
             'http://sdkhttp.eucp.b2m.cn/sdkproxy/ 'http://sdkhttp.eucp.b2m.cn/sdkproxy/regist.action
smsHttpHost = "sdkhttp.eucp.b2m.cn" 
smsXmlUser = "?cdkey="&smcUser&""&_
             "&password="&smcPass&""
smsXmlUEnc = Replace(smsXmlUser,smcPass,smcPase)
			 
			 
smsRetCode = "-1;0;304;305;307;308;3;10;11;12;13;14;15;16;17;18;22;27"
smsRetName = "未知错误;成功;客户端发送三次失败;服务器返回了错误的数据，原因可能是通讯过程中有数据丢失;发送短信目标号码不符合规则，手机号码必须是以0、1开头;非数字错误，修改密码时如果新密码不是数字那么会报308错误;连接过多，指单个节点要求同时建立的连接数过多;客户端注册失败;企业信息注册失败;查询余额失败;充值失败;手机转移失败;手机扩展转移失败;取消转移失败;发送信息失败;发送定时信息失败;注销失败;查询单条短信费用错误码"


'// 获取余额
function smsBalance()
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sFee
  aHttp = smpPost(smsHttpUrl,"querybalance","")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sFlag = smpXmlVal(aHttp(2),"error")
	If isNumeric(sFlag) Then
	  sFee = smpXmlVal(aHttp(2),"message")
	  If sFlag="0" Then
        rArr(0) = true
	    rArr(1) = sFee/0.1
        rArr(2) = sFee&":"&"正常!"
	  Else
	    rArr(1) = sFlag
        rArr(2) = sFlag&":"&smpState(smsRetCode,smsRetName,sFlag)
	  End If
	Else 
	    rArr(1) = -1
        rArr(2) = smpState(smsRetCode,smsRetName,"-1")
	End If
  Else
    rArr(1) = -1
	rArr(2) = smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsBalance = rArr
end function

'// 接口充值 
function smsCharge(xID,xPW,xMoney,xCount,xNote)
  Dim r,rArr(2),sql :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg :sSoap=""
  sSoap = sSoap& "&cardno="&xID
  sSoap = sSoap& "&cardpass="&xPW
  aHttp = smpPost(smsHttpUrl,"chargeup",sSoap)
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sFlag = smpXmlVal(aHttp(2),"error")
	If isNumeric(sFlag) Then
	  sFee = smpXmlVal(aHttp(2),"message")
	  If sFlag="0" Then
        rArr(0) = true
	    rArr(1) = sFee/0.1
        rArr(2) = "成功!"
		LogID = "FFFF"&Mid(Get_AutoID(24),5) '0BA94D4F-8FB3-25E1-WKS3H
		LogTime = DateAdd("yyyy",100,Now())
		sql =      "UPDATE [SmsCharge] SET "
		sql = sql& " LogID = '" & LogID &"'"
		sql = sql& ",LogMoney = LogMoney+" & xMoney &""
		sql = sql& ",LogCount = LogCount+" & xCount &""
		sql = sql& ",LogNote = '测试API:系统余额@" & Now() &"'"
		sql = sql& ",LogAddIP='"&Get_CIP()&"',LogAUser='"&Session("UsrID")&"',LogATime='"&LogTime&"' "
		sql = sql& " WHERE LogUser='Balance@TestAPI' "
		Call rs_DoSql(conn,sql)
	  Else
	    rArr(1) = sFlag
        rArr(2) = sFlag&":"&smpState(smsRetCode,smsRetName,sFlag)
	  End If
	Else 
	    rArr(1) = -1
        rArr(2) = smpState(smsRetCode,smsRetName,"-1")
	End If
  Else
    rArr(1) = -1
	rArr(2) = smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,smcDMod)
  smsCharge = rArr
end function

'// 发送短信
function smsSend(xNumb,xCont,xTime,xPort) 
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg,sAct :sSoap=""
  sSoap = sSoap& "&phone="&xNumb
  sSoap = sSoap& "&message="&Server.URLEncode(xCont)
  If(xTime<>"") Then
    sSoap = sSoap& "&sendtime="&smpFmtTime(xTime) 
	sAct = "sendtimesms"
  Else
    sAct = "sendsms"
  End If
  sSoap = sSoap& "&addserial="&xPort
  aHttp = smpPost(smsHttpUrl,sAct,sSoap)
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"error")
	If isNumeric(sState) Then
	  If Int(sState)=0 Then
        sMsg = smpXmlVal(aHttp(2),"message")
		rArr(0) = true
	    rArr(1) = sFee
        rArr(2) = "成功!"
	  Else
	    rArr(1) = sState 
        rArr(2) = sState&":"&smpState(smsRetCode,smsRetName,sState)
	  End If
	Else
	    rArr(1) = -1
        rArr(2) = smpState(smsRetCode,smsRetName,"-1")
	End If
  Else
    rArr(1) = -1
	rArr(2) = smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsSend = rArr
End Function


'// 测试发送
function smsTest(xNumb,xCont) 'SendSMS
  If xCont="" Then xCont="测试:"&Now()
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg :sSoap=""
  sSoap = sSoap& "&phone="&xNumb
  sSoap = sSoap& "&message="&Server.URLEncode(xCont)
  aHttp = smpPost(smsHttpUrl,"sendsms",sSoap)
  'Response.Write smpXmlShow(sSoap)
  Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"error")
    If cStr(sState)="0" Then
      sMsg = smpXmlVal(aHttp(2),"message")
	  rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "成功!"
	Else
	  rArr(1) = -1	  
	  rArr(2) = sState&":"&smpState(smsRetCode,smsRetName,sState)
    End If
  Else
    rArr(1) = -1
	rArr(2) = smpState(smsRetCode,smsRetName,"-1")
  End If
  Call smpDebug(rArr,aHttp(2))
  smsTest = rArr
END function


'// 修改密码
function smsPass(xPass,xPNew,xExt) 'UDPPwd
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg :sSoap=""
  sSoap = sSoap& "&newPassword="&xPNew
  aHttp = smpPost(smsHttpUrl,"changepassword",sSoap)
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"error")
    If cStr(sState)="0" Then
      sMsg = smpXmlVal(aHttp(2),"message")
	  rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "成功!"
	Else
	  rArr(1) = -1	  
	  rArr(2) = sState&":"&smpState(smsRetCode,smsRetName,sState)
    End If
  Else
    rArr(1) = -1
	rArr(2) = smpState(smsRetCode,smsRetName,"-1")
  End If
  Call smpDebug(rArr,aHttp(2))
  smsPass = rArr
END function

'// 接收短信
function smsReceive(xSub,xExt) 'mo/RECSMS
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg :sSoap=""
  aHttp = smpPost(smsHttpUrl,"getmo",sSoap)
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"error")
    If cStr(sState)="0" Then
      sMsg = smpXmlVal(aHttp(2),"message")
	  rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "成功!"&sMsg
	Else
	  rArr(1) = -1	  
	  rArr(2) = sState&":"&smpState(smsRetCode,smsRetName,sState)
    End If
  Else
    rArr(1) = -1
	rArr(2) = smpState(smsRetCode,smsRetName,"-1")
  End If
  Call smpDebug(rArr,aHttp(2))
  smsReceive = rArr
END function


'// 取回执
function smsReport(xLast,xCount,xExt) 'report
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState
  sSoap = ""&_
  "<report xmlns=""http://tempuri.org/"">"&_
    smsXmlUEnc&_ 
	"<maxid>"&xLast&"</maxid>"&_
  "</report>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"report")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"reportResult")
    If cStr(sState)="0" Then
      rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "OK"
	Else
	  rArr(1) = -1	  
	  rArr(2) = sState
    End If
  Else
    rArr(1) = -1
	rArr(2) = smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsReport = rArr
END function


%>

