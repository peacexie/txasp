<!--#include file="../himg/tconfig.asp"-->
<!--#include file="oConfig.asp"-->

<%
sTimer = Timer()

UsrIP = Get_CIP()
PagRef = Request.Servervariables("HTTP_REFERER")&""
If PagRef<>"" Then
PagRef = Replace(PagRef,"http://","") 
PagRef = Mid(PagRef,inStr(PagRef,"/"))
End If

If Session("SesID")&""="" Then

  Session("SesID") = "SesID_"&Session.SessionID
  '所有浏览量+1
  Call rs_DoSql(conn,"UPDATE OnlView SET vNum=vNum+1 WHERE vID='AllView'")
  '在线记录
  If rs_Exist(conn,"SELECT pID FROM OnlPage WHERE pID='"&Session("SesID")&"'")="EOF" Then
  Call rs_DoSql(conn,"INSERT INTO OnlPage (pID,pIP,pPage,pTime) VALUES ('"&Session("SesID")&"','"&UsrIP&"','"&PagRef&"',"&cfgTimeC&""&Now()&""&cfgTimeC&")")
  End If
  'Check今日浏览
  If rs_Exist(conn,"SELECT vID FROM OnlView WHERE vTime="&cfgTimeC&""&Date()&""&cfgTimeC&" AND vID<'9999'")="EOF" Then
    Call rs_DoSql(conn,"INSERT INTO OnlView (vID,vNum,vTime) VALUES ('"&Date()&"',1,"&cfgTimeC&""&Date()&""&cfgTimeC&")")
  Else
    Call rs_DoSql(conn,"UPDATE OnlView SET vNum=vNum+1 WHERE vTime="&cfgTimeC&""&Date()&""&cfgTimeC&" AND vID<'9999'" )
  End If
  '最高在线人数
  OnlMax = rs_Val("","SELECT vNum FROM OnlView WHERE vID='OnlMax'")
  OnlNow = rs_Count(conn,"OnlPage WHERE pTime>="&cfgTimeC&""&DateAdd("n",(-1)*pExpTime,Now())&""&cfgTimeC&"")
  Application("OnlNow") = OnlNow
  Application("OnlMax") = OnlMax
  If Int(OnlMax)<Int(OnlNow) Then
    Call rs_DoSql(conn,"UPDATE OnlView SET vNum="&OnlNow&",vTime="&cfgTimeC&""&Now()&""&cfgTimeC&" WHERE vID='OnlMax'" )
	Application("OnlMax") = OnlNow
  End If
  

  '刷新文件   '//////////////////////////////////////////////////////////////
FlgMod = Request("FlgMod")&""
FlgHH = DatePart("h",Now())&"@"&Date()
Application("FlgMod") = Application("FlgMod")&""
If inStr(Application("FlgMod"),FlgHH)<=0 Then
  '删除旧数据;数据少,操作快,1小时执行一次;
  Call rs_DoSql(conn,"DELETE FROM OnlView WHERE vTime<"&cfgTimeC&""&DateAdd("d",(-1)*pOldDays,Date())&""&cfgTimeC&" AND vID<'9999'")
  Call rs_DoSql(conn,"DELETE FROM OnlPage WHERE pTime<"&cfgTimeC&""&DateAdd("n",(-1)*pExpTime,Now())&""&cfgTimeC&"")
  Application("FlgMod") = "("&FlgHH&"),"
End If
If inStr("{"&FlgMod&"}","Home")>0 Then 
  If inStr(Application("FlgMod"),FlgMod)<=0 Then   '///////////////////////////
    Application("FlgMod") = Application("FlgMod")&"("&FlgMod&"),"
    Set rs=Server.Createobject("Adodb.Recordset")
    Set fso = CreateObject("Scripting.FileSystemObject") 
    Call aLogs(FlgMod)

    Function aLogs(xAct)
      Dim ymd,sym,sno,fil,s
      ymd = Get_FmtID("yymd","") 
      s = Now()&" : "&xAct&" : "&CStr(Timer()-sTimer)&" : "&Application("FlgMod")
      fp = Config_Path&"upfile/#dbf#/Cnt"&ymd&".txt"
      If fExist(Config_Path&"upfile/#dbf#/","Cnt"&ymd&".txt") Then
	    Set fil2 = fso.OpenTextFile(Server.MapPath(fp),8,True)
	    fil2.WriteLine(s) 
      Else
        Set fil2 = fso.CreateTextFile(Server.MapPath(fp),True)
	    fil2.Write(s&vbcrlf)
      End If
      fil2.Close
    End Function
    Function fExist(xpth,xnam)	
     Dim fp2
     fp2=Server.MapPath(xpth&xnam)
	 fExist=FSO.FileExists(fp2) 
    End Function

    Set rs = Nothing
    Set fso = Nothing
  End If   '////////////////////////////////////////////////////////////////////
End If



Else

'//////////////////////////////////////
  '打开记录...操作
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM OnlPage WHERE pID='"&Session("SesID")&"'",conn,1,1 
  if rs.eof then 
'//////////////////////////////////////

    'Check超时被清除了;
	'所有浏览量,今日浏览量+1
    Call rs_DoSql(conn,"UPDATE OnlView SET vNum=vNum+1 WHERE vID='AllView'")
	Call rs_DoSql(conn,"UPDATE OnlView SET vNum=vNum+1 WHERE vTime="&cfgTimeC&""&Date()&""&cfgTimeC&" AND vID<'9999'" )
    '在线记录
    Call rs_DoSql(conn,"INSERT INTO OnlPage (pID,pIP,pPage,pTime) VALUES ('"&Session("SesID")&"','"&UsrIP&"','"&PagRef&"',"&cfgTimeC&""&Now()&""&cfgTimeC&")")

'//////////////////////////////////////
  else
  'Response.Write "alert('SesID=SesID');"
'//////////////////////////////////////

    '得到数据
   pIP = rs("pIP")
   pPage = rs("pPage")
   pTime = rs("pTime")
   nMins = Abs( DateDiff("n",pTime,Now()) )
   nPage = inStr(pPage,PagRef)
   If nPage<=0 Then
     pPage = pPage&";"&vbcrlf&PagRef
	 pPage = Right(pPage,3200)
   End If
   If Int(nMins)>Int(pExpTime) Or nPage<=0 Then '超时or切换页;
   'Response.Write "alert(' > OR Page ');"
	  '所有浏览量,今日浏览量+1
      Call rs_DoSql(conn,"UPDATE OnlView SET vNum=vNum+1 WHERE vID='AllView'")
	  Call rs_DoSql(conn,"UPDATE OnlView SET vNum=vNum+1 WHERE vTime="&cfgTimeC&""&Date()&""&cfgTimeC&" AND vID<'9999'" )
	  Call rs_DoSql(conn,"UPDATE OnlPage SET pPage='"&pPage&"',pTime="&cfgTimeC&""&Now()&""&cfgTimeC&" WHERE pID='"&Session("SesID")&"' " )
   Else '未超时,更新时间;
      Call rs_DoSql(conn,"UPDATE OnlPage SET pTime="&cfgTimeC&""&Now()&""&cfgTimeC&" WHERE pID='"&Session("SesID")&"' " )
   End If

'//////////////////////////////////////
  end if 
rs.Close()
SET rs=Nothing 
'//////////////////////////////////////
 

End If


If Request("send")="View" Then

  Response.Write "<br>UsrIP:"&UsrIP
  Response.Write "<br>PagRef:"&PagRef
  Response.Write "<br>SesID:"&Session.SessionID
  Response.Write "<br>FlgMod:"&FlgMod
  Response.Write "<br>"&MaxView
  
  Response.Write "<br>FlgHH:"&FlgHH
  Response.Write "<br>sApp:"&Application("FlgMod")
  Response.Write "<br>vLogs:<a target='_blank' href='"&Config_Path&"upfile/%23dbf%23/Cnt"&Get_FmtID("yymd","")&".txt'>Logs</a>"
  
ElseIf Request("send")="Clear" Then 

  Application("FlgMod") = ""
  Session("SesID") = ""
  Application("OnlNow") = ""
  Application("OnlMax") = ""
  
ElseIf Request("send")="sCount" Then 

  cNum = rs_Val("","SELECT vNum FROM OnlView WHERE vID='AllView'")
  sNum = "" ': Response.Write sAllView
  If Config_NStyle="00" Then
    sNum = cNum
  Else
    For i=1 To Len(cNum)
     c = Mid(cNum,i,1)
     sNum = sNum&"<img src=\'"&Config_Path&"img/inum/"&Config_NStyle&"0"&c&".gif\' align=absmiddle border=0>"
    Next
  End If
  Response.Write "document.write('"&sNum&"');"

End If

%>

