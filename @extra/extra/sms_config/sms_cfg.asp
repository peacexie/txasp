<%

'说明: 此文件函数，为smc配置文件，
'所有 变量/函数 以smc开头


Dim smcMClass,smcMObj 'Module:组件名,组件对象 
Dim smcUser,smcUEID,smcPass,smcPase '用户名/序列号,企业ID,密码,加密码
Dim smcTSplit,smcTMax1,smcTMax2 'Tel:群发号码数,号码分割号
Dim smcCSplit,smcCLong,smcCMul 'Content:内容分隔号,长内容支持,多条内容支持
Dim smcDebug,smcDConn,smcDMod '调试标记,数据库连接,当前组件标记
Dim smcECfg : smcECfg = "20100424" '简易加密/解密密匙
Dim smcDTabs '不支持函数列表 'Charge;Pass;Receive;Report
Dim smcSTime,smcSPort '子端口，定时发送
Dim smcOTMax : smcOTMax = 20 '外部tel群发号码数

smcCSplit = "(mS9)" ' 内容分隔号
smcDConn = conn ' 数据库连接
'smcDebug = "" '调试标记: 见sms_cfg.asp

%>

