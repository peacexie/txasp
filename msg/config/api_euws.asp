<%

'<!--/*  sms_cfg.asp 定义的配置
'亿美软通 接口 
smcMClass = "" 'EUCPComm.dll
smcUser = "0SDK-EMY-0130-LH??P" ' 
'smcUEID = "" ' 企业ID
smcPass = "20110??5" '密码 txedu
smcPase = "145??6" '
smcTSplit = "," ' 号码分割号
smcTMax1 = 1000'3,1000 ' 接口群发号码数
smcTMax2 = 3200'4,9600 ' 系统群发号码
smcCLong = false ' true,false长内容
smcCMul = false ' true,false多条内容
smcDMod = "emay.cn亿美(ws)"
smcDTabs = "(Report;)" 'Charge;Receive;Report
smcSTime = true ' true,false定时发送
smcSPort = false ' true,false子端口
smcSFile = "api_euws.asp"
'*/-->

'说明: 此文件为 调用接口的函数 
'不同接口使用不同函数， smsSend()函数名不变，内容根据需要改变；
'所有 变量/函数 以sms开头

'接口公共数据 
smsHttpUrl = "http://sdkhttp.eucp.b2m.cn/sdk/SDKService?wsdl" 
             'http://117.79.237.3:8060/webservice.asmx '
smsHttpHost = "sdkhttp.eucp.b2m.cn" '219.239.91.112,sdkhttp.eucp.b2m.cn 
smsXmlUser = "<softwareSerialNo>"&smcUser&"</softwareSerialNo>"&_
             "<key>"&smcPass&"</key>"
smsXmlUEnc = Replace(smsXmlUser,smcPass,smcPase)

			 
smsRetCode = "0;10;101;305;999;-1;11;307;22;13;17;18;997;998;308"
smsRetName = "操作成功;客户端注册失败;客户端网络故障;服务器端返回错误(返回值不是数字字符串);操作频繁;企业信息或密码不符合要求;企业信息注册失败;目标电话号码不符合规则(01开头);注销失败;充值失败;发送信息失败;发送定时信息失败;找不到超时的短信无法确定是否成功;客户端网络问题导致发送超时;新密码不是数字"


'// 获取余额
function smsBalance()
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sFee
  sSoap = ""&_
  "<getBalance xmlns=""http://sdkhttp.eucp.b2m.cn/"">"&_ 
    smsXmlUser&_
  "</getBalance>"
  ''tempuri.org,sdkhttp.eucp.b2m.cn
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"getBalance")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sFee = smpXmlVal(aHttp(2),"return")
	If isNumeric(sFee) Then
	  If Int(sFee)>=0 Then
        rArr(0) = true
	    rArr(1) = sFee
        rArr(2) = "正常!"
	  Else
	    rArr(1) = sFee
        rArr(2) = sFee&":"&smpState(smsRetCode,smsRetName,sFee)
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

'// 发送短信
function smsSend(xNumb,xCont,xTime,xPort) 
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg,sAct :sSoap=""
  sSoap = ""&_
  "<sendSMS xmlns=""http://sdkhttp.eucp.b2m.cn/"">"&_
    smsXmlUser&_
    "<mobiles>"&smpFmtTime(xNumb)&"</mobiles>"&_
    "<smsContent>"&Server.URLEncode(xCont)&"</smsContent>"&_
    "<sendTime>"&smpFmtTime(xTime)&"</sendTime>"&_
    "<addSerial>"&xPort&"</addSerial>"&_
	"<srcCharset>UTF-8</srcCharset>"&_
	"<smsPriority>3</smsPriority>"&_
  "</sendSMS>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"sendSMS")  
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"return")
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
  sSoap = ""&_
  "<sendSMS xmlns=""http://sdkhttp.eucp.b2m.cn/"">"&_
    smsXmlUser&_
    "<mobiles>"&smpFmtTime(xNumb)&"</mobiles>"&_
    "<smsContent>"&Server.URLEncode(xCont)&"</smsContent>"&_
	"<srcCharset>UTF-8</srcCharset>"&_
	"<smsPriority>3</smsPriority>"&_
  "</sendSMS>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"sendSMS") 
  'Response.Write smpXmlShow(sSoap)
  Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"return")
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

'// 激活序列号
function smsRegID() '
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg :sSoap=""
  sSoap = ""&_
  "<registEx xmlns=""http://sdkhttp.eucp.b2m.cn/"">"&_
    "<softwareSerialNo>"&smcUser&"</softwareSerialNo>"&_
    "<key>"&smcPass&"</key>"&_
    "<serialpass>"&smcPase&"</serialpass>"&_
  "</registEx>" 
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"registEx") 
  'Response.Write "<br>("&smpXmlShow(sSoap)&")<br>"
  smcDebug = "isDebug"
  Call smpDebug(rArr,aHttp(2))
  smsRegID = rArr
END function

'// 注销序列号
function smsRegUn() '
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sMsg :sSoap=""
  sSoap = ""&_
  "<logout xmlns=""http://tempuri.org/"">"&_
    "<softwareSerialNo>"&smcUser&"</softwareSerialNo>"&_
    "<key>"&smcPass&"</key>"&_
  "</logout>" '187740
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"logout") 
  Response.Write smpXmlShow(sSoap)
  smcDebug = "isDebug"
  Call smpDebug(rArr,aHttp(2))
  smsRegUn = rArr
END function

%>

