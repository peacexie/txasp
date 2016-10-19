<%

'<!--/*  sms_cfg.asp 定义的配置
'小灵通短信网上辅助发送系统接口 - 张长江提供
smcMClass = "ZipCOM.TSecure_zip" '
smcUser = "dgtest1" ' 用户名/序列号
'smcUEID = "" ' 企业ID
smcPass = "abc??8" ' 密码
smcPase = "e10adc3949ba59abbe56e057f20f8??e"
smcTSplit = "," ' 号码分割号
smcTMax1 = 2000'3,1000 ' 接口群发号码数
smcTMax2 = 6400'4,9600 ' 系统群发号码数
smcCLong = false ' true,false长内容
smcCMul = true ' true,false多条内容
smcDMod = "noName张长江提供"
smcDTabs = "(Charge;Pass;Receive)" 'Charge;Receive;Report
smcSTime = true ' true,false定时发送
smcSPort = false ' true,false子端口
smcSFile = "api_chch.asp"
'*/-->

'说明: 此文件为 调用接口的函数 
'不同接口使用不同函数， smsSend()函数名不变，内容根据需要改变；
'所有 变量/函数 以sms开头

'接口公共数据
smsHttpUrl = "http://61.145.168.234:90/Interface.asmx" 
             'http://61.145.168.234:90/SMS_Interface.asmx
smsHttpHost = "61.145.168.234" 
smsXmlUser = "<UserName>"&smcUser&"</UserName>"&_
             "<UserPwd>"&smcPase&"</UserPwd>"
smsXmlUEnc = Replace(smsXmlUser,smcPass,smcPase)


smsXmlMsgs = "<Ms c=""2""><m>(mArr)</m></Ms>"


function smsSend(xNumb,xCont,xTime,xPort) 
  Dim sMsg,iMsg,i,iMax,aMsg
  aMsg = Split(xCont,smcCSplit) ':Response.Write xCont&smcCSplit
  sMsg = "" : iMax = uBound(aMsg)
  For i=0 To iMax
    iMsg = "<m>"&_
	        "<FD>"&xNumb&"</FD>"&_
	        "<FM>"&aMsg(i)&"</FM>"
    If xPort<>"" Then 
      iMsg = iMsg&"<FO>"&xPort&"</FO>"
    End If
    If xTime<>"" Then 
      iMsg = iMsg&"<FT>"&xTime&"</FT>"
    End If
    iMsg = iMsg&"</m>"
    sMsg = sMsg&iMsg 
  Next
  sMsg = Replace(smsXmlMsgs,"<m>(mArr)</m>",sMsg)
  'Response.Write smpXmlShow(sMsg)
  'Response.End()
  sMsg = smsCompress(sMsg) 
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sFee
  sSoap = ""&_
  "<Sms_SendEx2 xmlns=""http://tempuri.org/"">"&_
    "<CompCode></CompCode>"&_
    smsXmlUser&_
    "<SendMsgXML>"&sMsg&"</SendMsgXML>"&_
    "<withfollow>0</withfollow>"&_
  "</Sms_SendEx2>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"Sms_SendEx2")
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"Sms_SendEx2Result")
    If cStr(sState)="0" Then
      rArr(0) = true
	  rArr(1) = 0
      rArr(2) = "成功!"&smsdeComp(smpXmlVal(aHttp(2),"retValueStr"))
	Else
	  rArr(1) = -1	  
	  rArr(2) = smpState("-1;-2;-3;-4;-5;-9","主叫或被叫号码有误;内容含非法关键字;账号或密码有误;存在禁发号码;超过账号发送额度;异常",aHttp(0))
    End If
  Else
    rArr(1) = -1
	rArr(2) = aHttp(0)&" - "&aHttp(1)&" - "&smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsSend = rArr
End Function

function smsTest(xNumb,xCont)
  If xCont="" Then xCont="测试:"&Now()
  Dim sNumb,sCont,sMsg
  sNumb = "<FD>"&xNumb&"</FD>"
  sCont = "<FM>"&xCont&"</FM>"
  sMsg = Replace(smsXmlMsgs,"(mArr)",sNumb&sCont)
  'Response.Write smpXmlShow(sMsg)
  sMsg = smsCompress(sMsg)
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sFee
  sSoap = ""&_
  "<Sms_SendEx xmlns=""http://tempuri.org/"">"&_
    "<CompCode></CompCode>"&_
    smsXmlUser&_
    "<SendMsgXML>"&sMsg&"</SendMsgXML>"&_
    "<withfollow>0</withfollow>"&_
  "</Sms_SendEx>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"Sms_SendEx")
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"Sms_SendExResult")
    If cStr(sState)="0" Then
      rArr(0) = true
	  rArr(1) = 1
      rArr(2) = "成功!"&smsdeComp(smpXmlVal(aHttp(2),"retValueStr"))
	Else
	  rArr(1) = 0	  
	  rArr(2) = smpState("-1;-2;-3;-4;-5;-9","主叫或被叫号码有误;内容含非法关键字;账号或密码有误;存在禁发号码;超过账号发送额度;异常",aHttp(0))
    End If
  Else
    rArr(1) = 0
	rArr(2) = aHttp(0)&" - "&aHttp(1)&" - "&smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsTest = rArr
End Function

'// 获取余额
'Dim rState,rSText,rText,oHttp
function smsBalance() 
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sFee
  sSoap = ""&_
  "<GetMeasure xmlns=""http://tempuri.org/"">"&_
    smsXmlUser&_
	"<type>0</type>"&_
  "</GetMeasure>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"GetMeasure")
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(aHttp(2),"GetMeasureResult")
    If cStr(sState)="0" Then
      rArr(0) = true
	  sFee = smpXmlVal(aHttp(2),"cMount")
	  rArr(1) = sFee
      rArr(2) = ""
	Else
	  rArr(1) = 0
	  rArr(2) = smpState("-1;-2;-3;-9","调用失败;余额数据异常;账号或密码有误;账号查询错误",sState)
    End If
  Else
    rArr(1) = 0
	rArr(2) = aHttp(0)&" - "&aHttp(1)&" - "&smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  smsBalance = rArr
End Function

'// 修改密码
function smsPass() 'UDPPwd
  '
END function

'// 取回执
function smsReport(xLast,xCount,xExt) 
  Dim rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  Dim sSoap,aHttp,sState,sFee
  sSoap = ""&_
  "<GetReport2 xmlns=""http://tempuri.org/"">"&_
    smsXmlUser&_
	"<lastIndex>"&xLast&"</lastIndex>"&_
	"<count>"&xCount&"</count>"&_
	"<type>"&xExt&"</type>"&_
  "</GetReport2>"
  aHttp = smpHttp(smsHttpUrl,smsHttpHost,sSoap,"GetReport2")  
  'Response.Write smpXmlShow(sSoap)
  'Response.Write smpXmlShow(aHttp(2))
  If cStr(aHttp(0))="200" Then
    sState = smpXmlVal(oHttp,"GetReport2Result")
	sState = smsdeComp(sState) 
	rArr(0) = true
	rArr(1) = 0
	rArr(2) = "成功"&"(sState:"&sState&":End)"
  Else
    rArr(1) = 0
	rArr(2) = aHttp(0)&" - "&aHttp(1)&" - "&smpXmlText(aHttp(2))
  End If
  Call smpDebug(rArr,aHttp(2))
  Set oHttp = Nothing
  smsReport = rArr
End Function

'// 压缩Msg(Xml)
function smsCompress(sXml) 
  smpOSet()
     smsCompress = smcMObj.Compress(sXml)
  smpOEnd() 
End Function
'// 解压Msg(Xml)
function smsdeComp(sStr) 
  smpOSet()
     smsdeComp = smcMObj.deCompress(sStr)
  smpOEnd() 
End Function

'// 调试状态：smcDebug = "Debug"
Function smpDebug(rArr,rText) 
  If smcDebug = "isDebug" Then
	Response.Write "<br>"&rArr(0)
	Response.Write "<br>"&rArr(1)
	Response.Write "<br>"&rArr(2)
	Response.Write "<hr>"&smpXmlShow(rText)
  End If
End Function

  '返回-1：调用失败,-2余额数据异常，-9账号查询错误，0成功
  '0成功；-1 主叫或被叫号码有误；-2:内容含非法关键字；-3账号或密码有误；-4：存在禁发号码；-5:超过账号发送额度，-9异常


'3.11  Register  注册	8
'3.12  ChargUp 充值	8
'3.13  GetBalance  获取余额	9
'3.14  SendSMS  发送短信	9
'3.15  SendSMSEx  发送短信(网络版)	9
'3.16  RECSMS 接收短信	9
'3.17  RECSMSEx接收短信(网络版)	10
'3.18  UDPPwd 更改密码	10
'3.21  UnRegister 注销	10

%>

