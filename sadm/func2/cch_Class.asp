<!--#include file="cch_Config.asp"-->
<%

Function cchCheck(cfPath)
  Dim url,q
  If Left(Request("~act"),9)="~getHttp." Then
    Exit Function 
  Else
    If cchFlag<>"Null" Then ' ------------------ 1 cchFlag = RAM/File
      url = Request.ServerVariables("URL")
      q = Request.QueryString()
	  If q<>""     Then url = url&"?"&q
      If cfPath="" Then cfPath = url
      If inStr(cchTabs,cfPath) Then ' ---------- 3 在列表范围内
        If cchFlag="File" Then
          Call cchFile(cfPath)
		Else
          Call cchRAM(cfPath)
        End If
      End If
    End If 
  End If
End Function
Function cchRAM(cfPath)
    Dim myCache,myName,flag,chCont
    myName = "Cache["&strFill(cfPath)&"].32."
    set myCache = new clsCache
      myCache.name = myName 
      if myCache.valid then 
        chCont = myCache.value 
        flag = "Read"
      else
        chCont = urlRead(cfPath&"?~act=~getHttp."&strRand()) 
        myCache.add chCont,dateadd("n",cchTime,now) 
        flag = "Upd"
      end if 
    set myCache = nothing 
    Call strShow(chCont,flag,"cchRAM")
End Function
Function cchFile(cfPath)
    Dim cfName,cfFull,chObj,flag,chCont
    flag = "Null" : chCont = ""
    cfName = "Cache["&strFill(cfPath)&"].31.htm" 
    cfFull = Config_Path&"upfile/sys/cache/"&cfName
    chObj = application(cfName) 
    if isempty(chObj) or (not isdate(chObj)) then ' 空,不是日期,更新
      flag = "Upd"
    elseif CDate(chObj)<now then                  ' 日期过期,更新
      flag = "Upd"                                
    else 
      flag = "Read"
    end if
    if flag = "Read" then                        
        'chCont = fileRead(cfFull,"utf-8")         ' 有效:读文件 '另判断文件是读文件的约1/5
		Server.Execute(cfFull)                     ' 用时是读文件的1/10~1/45
    else                                         
        application.lock                          ' 更新application
        application(cfName) = dateadd("n",cchTime,now)
        application.unlock
        chCont = urlRead(cfPath&"?~act=~getHttp."&strRand())
        Call fileSave(cfFull,chCont,"utf-8")      ' 取数据,写文件
    end if 
    Call strShow(chCont,flag,"cchFile")
End Function
Function cchClear()
  'application.lock
  'For Each Key in Application.Contents
    ''application(Key) = empty
  'Next
  'application.unlock
  Application.Contents.Removeall()
  ''Application.Clear() 'XXX
  Dim iFrm,sMod,sTab,i
  iFrm = "IF"&"RAME"
  sDiv = "<span style='width:20px; height:10px; overflow:hidden; display:inline-block; border:1px solid #F0F0F0; margin:1px; '>(iFrm)</span>"
  sMod = "<"&iFrm&" src='(url)' frameBorder=0 width='60' height='20' scrolling='no'></"&iFrm&">"
  sMod = Replace(sDiv,"(iFrm)",sMod)
  aTab = Split(cchTabh,"|")
  For i=0 To uBound(aTab)
    If aTab(i)<>"" Then
	  Response.Write Replace(sMod,"(url)",aTab(i))
	End If
  Next
End Function


Function strShow(c,f,m)
    Response.Write c
    Response.Write vbcrlf&"<!-- "&f&"."&m&"-->"
    Response.End()
End Function
Function strFill(cfPath)
    Dim cfName
    cfName = Replace(cfPath,"/","~") ' page/index.asp -=> page~index.asp.cache
    cfName = Replace(cfName,"?","!") ' ?DepID=DepTech -=> !DepID=DepTech
    cfName = Replace(cfName,":","") ' 
    cfName = Replace(cfName,"|","") ' 
    strFill = cfName
End Function
Function strRand()
    Dim sNow '/:   
    sNow = Replace(Now()," ","")
    sNow = Replace(sNow,"-","")
    sNow = Replace(sNow,"/","")
    sNow = Replace(sNow,":","")
    Randomize 
    strRand = sNow&"."&Int(1234567890*Rnd())&"."&Int(1234567890*Rnd())
End Function


Function fileSave(xPName2,xCont,xCSet)
 Dim st,xPName : xPName=xPName2
 If instr(xPName,":")<=0 Then xPName=server.MapPath(xPName)  
 Set st=Server.CreateObject("ADODB.Stream")   
  st.Type=2   
  st.Mode=3   
  st.Charset=xCSet  
  st.Open()   
  st.WriteText xCont  
 st.SaveToFile xPName,2   
 st.Close()  
 Set st=Nothing   
End Function
Function fileExist(pathName)
  Dim fp : fp = Server.MapPath(pathName)
  SET FSO = Server.CreateObject("Scripting.FileSystemObject")
    fileExist = FSO.FileExists(fp) 
  Set FSO = Nothing
End Function
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


Function urlPath(uPath)
  Dim sPath,sUrl,sPort
  sPath = uPath
  If inStr(sPath,"://")<=0 Then
    sUrl = "http://"&Request.ServerVariables("SERVER_NAME")
    sPort = Request.ServerVariables("SERVER_PORT")
    If sPort<>"80" Then
      sUrl = sUrl&":"&sPort
    End If ' sUrl=http://www.domain.com:81
    If Left(sPath,1)<>"/" Then
      sPath = "/"&sPath
    End If ' page/index.asp -=> /page/index.asp
    If Len(sPath)-Len(Replace(sPath,"?",""))=2 Then
       sPath = Replace(sPath,"?~act=~","&~act=~")
    End If ' /page.asp?DepID=DepTech?~act=~getHttp.
    urlPath = sUrl&sPath 
  Else
    urlPath = sPath 
  End If
End Function
Function urlRead(strUrl) 
  Dim objHttp 
  strUrl = urlPath(strUrl)
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
  'objHttp.Charset = "gb2312"
  urlRead = objHttp.responseText  
  Set objHttp = Nothing 
End Function
Function urlSave(strUrl,strFile) 
  Dim Ads, objHttp, binData 
  If inStr(strUrl,"://")<=0 Then
    strUrl = urlPath(strUrl)
  End If
  On Error Resume Next 
  Set objHttp = Server.CreateObject("Micro"&"soft.XML"&"HTTP") 
  With objHttp 
    .Open "Get", strUrl, False, "", "" 
    .Send 
    binData = .ResponseBody 
  End With 
  Set objHttp = Nothing 
  Set Ads = Server.CreateObject("Ado"&"db.Str"&"eam") 
  With Ads 
    .Type = 1 
    .Open 
    .Write binData 
    .SaveToFile Server.MapPath(strFile), 2 
    .Cancel() 
    .Close() 
  End With 
  Set Ads=nothing 
  If err <> 0 then 
    urlSave = false 
    err.clear 
  Else 
    urlSave = true 
  End If 
End Function

%>


<% 
'-------------------------------------------------------------------------------------
'转发时请保留此声明信息,这段声明不并会影响你的速度!
'**************************   【先锋缓存类】Ver2004  ********************************
'作者:孙立宇、apollosun、ezhonghua
'官方网站:http://www.lkstar.com   技术支持论坛：http://bbs.lkstar.com
'电子邮件:kickball@netease.com    在线QQ:94294089
'版权声明:版权没有,盗版不究，源码公开,各种用途均可免费使用，欢迎你到技术论坛来寻求支持。
'目前的ASP网站有越来越庞大的趋势，资源和速度成了瓶颈
'——利用服务器端缓存技术可以很好解决这方面的矛盾，大大加速ASP运行和改善效率。
'本人潜心研究了各种算法，应该说，这个类是当前最快最纯净的缓存类。
'详细使用说明或范例请见下载附件或到本人官方站点下载！
'Peace 稍做修改 ...... 2010-05-17 14:30
'--------------------------------------------------------------------------------------

class clsCache

  private cchCont      '缓存内容
  private cchName      '缓存Application名称
  private cchTime      '缓存过期时间
  private cchTNam      '缓存过期时间Application名称
  private cchPath      '缓存页URL路径
  
  private sub class_initialize()
    cchPath=request.servervariables("url")
    cchPath=left(cchPath,instrRev(cchPath,"/"))
  end sub
      
  private sub class_terminate()
  end sub
  
  Public Property Get Version
      Version="先锋缓存类 Version 2004 [Edit by Peace]"
  End Property
  
  public property get valid '读取缓存是否有效/属性
    if isempty(cchCont) or (not isdate(cchTime)) then
      vaild=false
    elseif CDate(cchTime)<now then ' Peace //添加
      vaild=false                  ' Peace //添加
    else
      valid=true
    end if
  end property
  
  public property get value '读取当前缓存内容/属性
    if isempty(cchCont) or (not isDate(cchTime)) then
      value=null
    elseif CDate(cchTime)<now then
      value=null
    else
      value=cchCont
    end if
  end property
  
  public property let name(str) '设置缓存名称/属性
    cchName=str&cchPath&".main"
    cchCont=application(cchName)
    cchTNam=str&cchPath&".expire"
    cchTime=application(cchTNam)
  end property
  
  public property let expire(tm) '设置缓存过期时间/属性
    cchTime=tm
    application.Lock()
    application(cchTNam)=cchTime
    application.UnLock()
  end property
  
  public sub add(varCache,varExpireTime) '对缓存赋值/方法
    if isempty(varCache) or not isDate(varExpireTime) then
      exit sub
    end if
    cchCont=varCache
    cchTime=varExpireTime
    application.lock
    application(cchName)=cchCont
    application(cchTNam)=cchTime
    application.unlock
  end sub
  
  public sub clean() '释放缓存/方法
    application.lock
    application(cchName)=empty
    application(cchTNam)=empty
    application.unlock
    cchCont=empty
    cchTime=empty
  end sub
   
  public function verify(varcch2) '比较缓存值是否相同/方法——返回是或否
  if typename(cchCont)<>typename(varcch2) then
      verify=false
  elseif typename(cchCont)="Object" then
      if cchCont is varcch2 then
          verify=true
      else
          verify=false
      end if
  elseif typename(cchCont)="Variant()" then
      if join(cchCont,"^")=join(varcch2,"^") then
          verify=true
      else
          verify=false
      end if
  else
      if cchCont=varcch2 then
          verify=true
      else
          verify=false
      end if
  end if
  end Function
End Class
%>

