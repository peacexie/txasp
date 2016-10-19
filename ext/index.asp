<%

'/* Peace 说明 ***************************************

'默认, 每300秒 更新缓存一次；
'根据需要设置参数：/dghome.php:$cache_exp

'缓存管理 测试参数
'?is_upd=Clear 删除缓存,强制更新;  
'?is_upd=noUpd 不过期,直接读缓存文件; 快
'?is_upd=(其它不为空) 立即过期,强制更新缓存; 慢
 
'运行时间 测试参数
'?is_tst=(其它不为空)，显示读缓存文件的时间; 一般很小
'不设置，不显示; 
'（用于使用缓存下，显示读缓存文件的时间）

'* **************************************************/

tTimer = timer() '; //用于测试

cache_exp = 300 '; //缓存时间,单位:秒 (86400=24小时) 
cache_file = "./data/cache/dghome.htm" '; //缓存文件,[./]开头
real_file = "/peace/dgind.php" '; //要执行的真实文件,[/]开头
cache_flag = "Read" ';

Response.Write urlRead("http://b.dg.gd.cn/dghome.php") 

Function fileRead(pathName,xCSet)
 Dim str
 if fileExist(pathName) then
  set stm=server.CreateObject("Ado"&"db.Str"&"eam")
  stm.Type=2 
  stm.Mode=3 
  stm.Charset=xCSet
  stm.Open
  stm.loadfromfile Server.MapPath(pathName)
  str=stm.ReadText
  stm.Close
  set stm=nothing
 else
  str = "" '(File Read Error!)
 end if
 fileRead = str
End Function

Function urlRead(strUrl) 
  Dim objHttp 
  'strUrl = urlPath(strUrl)
  On Error Resume Next 
  Set objHttp = Server.CreateObject("Micro"&"soft.XML"&"HTTP") 
  With objHttp 
    .Open "Get", strUrl, False, "", "" 
    .Send 
  End With 
  if objHttp.Readystate <> 4 then 
    Set objHttp = Nothing 
    urlRead = False 
    Exit Function 
  end if 
  'objHttp.Charset = "utf-8"
  urlRead = objHttp.responseText  
  Set objHttp = Nothing 
End Function

%>

