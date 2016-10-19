<%

'<!--/*  sms_cfg.asp 定义的配置
'博星科技(众凯) 接口 
smcMClass = "" 'VBSDK.ASPCom
smcUser = "SDK-DHX-010-00011" ' 用户名/序列号 
'smcUEID = "" ' 企业ID
smcPass = "WHRX2T" '  ' 密码 343434,902130,WHRX2T
smcPase = "68A4026543695874F43244CDFB58149C" 'F09537114CC4736C69219F78FCD17EEC/68A4026543695874F43244CDFB58149C
smcTSplit = "," ' 号码分割号
smcTMax1 = 1000'3,1000 ' 接口群发号码数
smcTMax2 = 3200'4,9600 ' 系统群发号码数
smcCLong = true ' true,false长内容
smcCMul = false ' true,false多条内容
smcDMod = "bucp.net博星SDK"
smcDTabs = "(Charge;)" 'Charge;Receive;Report
smcSTime = true ' true,false定时发送
smcSPort = false ' true,false子端口
'*/-->

'说明: 此文件为 调用接口的函数 
'不同接口使用不同函数， smsSend()函数名不变，内容根据需要改变；
'所有 变量/函数 以sms开头

'接口公共数据
smsHttpUrl = "http://117.79.237.3:8060/webservice.asmx" 'http://sdk2.entinfo.cn/webservice.asmx
             'http://117.79.237.3:8060/webservice.asmx 'http://www.bucp.net/news/a0016.htm 'http://117.79.237.3:8060/webservice.asmx
smsHttpHost = "117.79.237.3" 'sdk2.entinfo.cn
smsXmlUser = "<sn>"&smcUser&"</sn>"&_
             "<pwd>"&smcPass&"</pwd>"
smsXmlUEnc = Replace(smsXmlUser,smcPass,smcPase)
			 
			 
smsRetCode = "0;-2;-3;-4;-5;-6;-7;-8;-9;-10;-11;-12;-13;-14;-15"
smsRetName = "成功;帐号/密码不正确;重复登陆;余额不足;数据格式错误;参数有误;权限受限;流量控制错误;扩展码权限错误;内容长度长;数据库错误;序列号状态错误;没有提交增值内容;服务器写文件失败;文件内容base64编码错误"


'// 获取余额
function smsBalance()
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sFee
  sSoap = ""&_
  "<GetBalance xmlns=""http://tempuri.org/"">"&_
    smsXmlUser&_
  "</GetBalance>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"GetBalance")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sFee = smpXmlVal(aHttp(2),"GetBalanceResult")
	If isNumeric(sFee) Then
	  If Int(sFee)>=0 Then
        rArr(0) = true
	    rArr(1) = sFee
        rArr(2) = "正常!"
	  Else
	    rArr(1) = sFee
        rArr(2) = "错误!"
	  End If
	Else
	    rArr(1) = -1
        rArr(2) = sFee
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
  Dim sSoap,aHttp,sState
  sSoap = ""&_
  "<mt xmlns=""http://tempuri.org/"">"&_
    smsXmlUEnc&_
    "<mobile>"&xNumb&"</mobile>"&_
    "<content>"&xCont&"</content>"&_
	"<ext>"&xPort&"</ext>"&_
	"<stime>"&xTime&"</stime>"&_
	"<rrid></rrid>"&_ 
  "</mt>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"mt")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"mtResult")
	If isNumeric(sState) Then
	  If Int(sState)>=0 Then
        rArr(0) = true
	    rArr(1) = sFee
        rArr(2) = "成功!"&smpXmlVal(aHttp(2),"mtResult")
	  Else
	    rArr(1) = sState
        rArr(2) = smpState(smsRetCode,smsRetName,sState)
	  End If
	Else
	    rArr(1) = -1
        rArr(2) = smpXmlText(aHttp(2))
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
  Dim sSoap,aHttp,sState
  sSoap = ""&_
  "<SendSMS xmlns=""http://tempuri.org/"">"&_
    smsXmlUser&_
    "<mobile>"&xNumb&"</mobile>"&_
    "<content>"&xCont&"</content>"&_ 
  "</SendSMS>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"SendSMS")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"SendSMSResult")
	'Response.Write "<br>200:"&sState
    If cStr(sState)="0 成功" Then
      rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "成功!"&smpXmlVal(aHttp(2),"mtResult")
	Else
	  rArr(1) = -1	  
	  rArr(2) = "失败！"
    End If
  Else
    rArr(1) = -1
	rArr(2) = smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsTest = rArr
END function


'// 修改密码
function smsPass(xPass,xPNew,xExt) 'UDPPwd
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState
  smsXmlUser = Replace(smsXmlUser,smcPass,xPass)
  sSoap = ""&_
  "<UDPPwd xmlns=""http://tempuri.org/"">"&_
    smsXmlUser&_
	"<newpwd>"&xPNew&"</newpwd>"&_
  "</UDPPwd>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"UDPPwd")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"UDPPwdResult")
	'Response.Write "<br>200:"&sState
    'If cStr(sState)="0 成功" Then
	If inStr(sState,"0")>0 And inStr(sState,"成功")>0 Then
      rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "成功!"
	Else
	  rArr(1) = -1	  
	  rArr(2) = sState
    End If
  Else
    rArr(1) = -1
	rArr(2) = smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsPass = rArr
END function

'// 接收短信
function smsReceive(xSub,xExt) 'mo/RECSMS
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState
  sSoap = ""&_
  "<RECSMSEx_UTF8 xmlns=""http://tempuri.org/"">"&_
    smsXmlUser&_
	"<subcode>"&xPNew&"</subcode>"&_
  "</RECSMSEx_UTF8>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"RECSMSEx_UTF8")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"RECSMSEx_UTF8Result")
	sState = smpXmlShow(sState)
	'Response.Write "<br>200:"&sState
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

