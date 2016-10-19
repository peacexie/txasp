<%Response.CodePage=65001%>
<%Response.Charset="utf-8"%>

<%

'Response.Write Config_Path
'Config_Path = "/" '配置好参数 Config_Path

Dim cchFlag : cchFlag = "RAM" 'RAM,File,Null // False,True
Dim cchTime : cchTime = 15 ' 分钟(10-60) 24*6=144

Dim cchTabh,cchTabs 

cchTabh =         Config_Path&"page/index.asp|"
cchTabh = cchTabh&Config_Path&"peng/index.asp|"

cchTabs = cchTabh
cchTabs = cchTabs&Config_Path&"ext/link.asp|"
'cchTabs = cchTabs&Config_Path&"page/dep.asp?DepID=DepTech|"


'Call cchCheck("") ''配置以上参数
'Call cchFile(Config_Path&"ext/map.asp")
'Call cchRAM(Config_Path&"ext/link.asp")

' Call fileSave(xPName2,xCont,xCSet)
' f = fileExist(pathName)
' s = fileRead(pathName,xCSet)
' url = urlPath(sPath)


%>

