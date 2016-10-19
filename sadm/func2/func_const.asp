<!--#include file="../../upfile/sys/config/Config.asp"-->
<%
' /upfile/sys/config/Config.asp 文件中定义配置信息 virtual
'12345678123456781234567812345678
'!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~  // + 空格,127
'!#$%&()*+,-./:;<=>?@[\]^_`{|}~    // "'

'//================= 资料存放配置 可无限扩展！
'$Config_upRoot = $_SERVER['DOCUMENT_ROOT'].'/upfile/';   //跟目录用此设置
'$Config_upRoot = 'F:/php/webs/vir_dirs/upfile/';         //虚拟目录用此设置
Dim Config_dbTab, Config_upDir, Config_mdKey
Config_dbTab = "InfoNews|InfoPics|BBSInfo|DocsNews|TradeInfo|GboInfo|GboSend|T04Info"    '41资料存放表格
Config_upDir = "dtinf|dtpic|dtbbs|dtdoc|dtbus|defdt|defdt|dt004"    '42附件存放目录
Config_mdKey = "Inf|Pic|BBS|Doc|Tra|Gbo|Gbo|M04"    '43模块标识前缀
'//================= MSSQL配置
Dim cfgDBType,cfgTimeC,cfgSqlServer,cfgSqlUser,cfgSqlPassword,cfgSqlDatabase
'cfgDBType = "MSSQL" '//必要，与下面conn同时设置
'cfgTimeC = "'" '时间比较
cfgSqlServer = "(local)"   'sql服务器   
cfgSqlUser = "sa"       '用户名   
cfgSqlPassword = "pass1234"     '密码    
cfgSqlDatabase = "ysWeb_0BA8F932_F1EE"
'//================= Access配置
cfgDBType = "Access" '//必要，与下面conn同时设置
cfgTimeC = "#" '时间比较 

Const Adm_aUser = "_12345"
Const Mem_aMemb = "_ABCDE"
Const conDriver = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="
Dim conn,conExt,conOnline,UsrPStr,MemPStr
UsrPStr="UsrP"&Config_Code
MemPStr="MemP"&Config_Code
'conn     = "Provider=SqlOleDB.1;Data Source="&cfgSqlServer&";UID="&cfgSqlUser&";PWD="&cfgSqlPassword&";Initial Catalog="&cfgSqlDatabase&""
conn      = conDriver&Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_"&Config_Code&".Peace!DB" ) 
conExt    = conDriver&Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_ExtData.Peace!DB" )
conOnline = conn 'conDriver&Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_Online.Peace!DB" )
'conOracle = "Driver={Microsoft ODBC for oracle};Server=mtc666;Uid=mtc666;Pwd=mypass;"
'cbbs     = conDriver&Server.MapPath(CDataBBS)
'connBak  = conn '(备份一份)
'Response.Write conn 


'Call SubDirect() ’建议只在前台使用
Function SubDirect() ' 域名跳转
  '忽略: 有端口,localhost,IP访问,HTTPS
  Dim host,dSub,url,q,dom1,dir1,dom2,dir2,h_cn,h_dot
  host = Request.ServerVariables("HTTP_HOST") 'localhost:240, 127.0.0.1:240, 
  If inStr(host,":")>0 Then '有端口，忽略
	Exit Function
  End If
  If inStr(host,"localhost")>0 Then 'localhost，忽略
	Exit Function
  End If
  If Request.ServerVariables("LOCAL_ADDR")=Request.ServerVariables("SERVER_NAME") Then 'IP访问，忽略
	Exit Function
  End If
  If inStr(Request.ServerVariables("SERVER_PROTOCOL"),"HTTPS")>0 Then 'HTTPS，忽略
	Exit Function
  End If
  '+www
  url = Request.ServerVariables("URL")
  q = Request.QueryString() :If q<>"" Then url = url&"?"&q
  dSub = "" '"dg.gd.cn" '固定
  If dSub="" Then '只处理通用domain.com, domain.com.cn
    If inStr(host,"www.")<=0 Then
	  h_cn = Replace(host,".cn","")
	  h_dot = Replace(h_cn,".","")
	  If Len(h_cn)-Len(h_dot)<=1 Then
	    Response.Redirect "http://www."&host&url
	  End If
	End If
  Else '处理特定
    If host=dSub Then
	  Response.Redirect "http://www."&host&url
	End If
  End If
  '不同域名跳转到目录 (蓝欣实业)
  dom1 = "domain~1.com" :dir1 ="/page/" 
  dom2 = "domain~2.net" :dir2 ="/peng/"
  url = lCase(url)
  If inStr(host,dom1)>0 And inStr(url,dir2)>0 Then
    Response.Redirect "http://www."&dom2&url
  End If
  If inStr(host,dom2)>0 And inStr(url,dir1)>0 Then
    Response.Redirect "http://www."&dom1&url
  End If
  '子域名跳转到目录: 见 /ext/go.asp
  'blog.domain.com -=> www.domain.com/blog/
End Function

%>