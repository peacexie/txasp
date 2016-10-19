<%

'<!--/*  sms_cfg.asp 定义的配置
'Tester 接口 
smcMClass = "Test.Class" '
smcUser = "0BA9484F-6498-00CA-1B53-PEACE0ASP013" ' 用户名/序列号 
'smcUEID = "" ' 企业ID
smcPass = "myPass" '  ' 密码
smcPase = "encPass" 
smcTSplit = "," ' 号码分割号
smcTMax1 = 3'3,1000 ' 接口群发号码数
smcTMax2 = 5'4,9600 ' 系统群发号码
smcCLong = true ' true,false长内容
smcCMul = false'true'false' ' true,false多条内容
smcDMod = "test.Peace"
smcDTabs = "(Pass;Receive;Report)" 'Charge;Pass;Receive;Report
smcSTime = false
smcSPort = true
'*/-->

'说明: 此文件为 调用接口的函数 
'不同接口使用不同函数， smsSend()函数名不变，内容根据需要改变；
'所有 变量/函数 以sms开头

'rArr(1)返回: 不是数字:-1,0; 是数字:-1,0,N

'// 获取余额
function smsBalance()
  Dim r,rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  r = 40123 + Rnd_ID("0",4)
  If r Mod 6 = 0 Then
    rArr(0) = false
    rArr(1) = -1
    rArr(2) = "失败!"
  Else
    rArr(0) = true
    rArr(1) = rs_Val(conn,"SELECT LogCount FROM [SmsCharge] WHERE LogUser='Balance@TestAPI' ")
    rArr(2) = "成功!"
  End If
  Call smpDebug(rArr,smcDMod)
  smsBalance = rArr
end function

'// 接口充值 
function smsCharge(xID,xPW,xMoney,xCount,xNote)
  Dim r,rArr(2),sql :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  r = 40123 + Rnd_ID("0",4)
  If r Mod 6 = 0 Then
    rArr(0) = false
    rArr(1) = -1
    rArr(2) = "失败!"
  Else
    rArr(0) = true
    rArr(1) = 0
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
  End If
  Call smpDebug(rArr,smcDMod)
  smsCharge = rArr
end function

'// 发送短信
function smsSend(xNumb,xCont,xTime,xPort) 
  Dim r,rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  r = 40123 + Rnd_ID("0",4)
  If r Mod 6 = 0 Then
    rArr(0) = false
    rArr(1) = -1
    rArr(2) = "失败!"
  Else
    rArr(0) = true
    rArr(1) = 0
    rArr(2) = "成功!"&Get_GUID("","")
	Call smsSubtract(xNumb,xCont)
  End If
  Call smpDebug(rArr,smcDMod)
  smsSend = rArr
End Function

'// 测试发送
function smsTest(xNumb,xCont) 
  If xCont="" Then xCont="测试:"&Now()
  Dim r,rArr(2) :rArr(0)=false :rArr(1)=0 :rArr(2)=""
  r = 40123 + Rnd_ID("0",4)
  If r Mod 6 = 0 Then
    rArr(0) = false
    rArr(1) = -1
    rArr(2) = "失败!"
  Else
    rArr(0) = true
    rArr(1) = 0
    rArr(2) = "成功!"&Get_GUID("","")
	Call smsSubtract(xNumb,xCont)
  End If
  Call smpDebug(rArr,rArr(2))
  smsTest = rArr
END function

'// 修改密码
function smsPass() '
  '
END function

'// 接收短信
function smsReceive() '
  '
END function

'// 取回执
function smsReport() 
  '
End Function

'Test系统扣费 'Subtract,Charge
function smsSubtract(xNumb,xCont)
  Dim nCnt,aNumb,nSend,sql
  nCnt = smpCntCont(xCont) 
  aNumb = Split(smpFmtTNum(xNumb),";") '格式化
  nSend = (uBound(aNumb)+1)*nCnt
  sql = "UPDATE [SmsCharge] SET LogCount = LogCount-" &nSend&" WHERE LogUser='Balance@TestAPI' "
  Call rs_DoSql(conn,sql) ':Response.Write sql
End Function

%>

