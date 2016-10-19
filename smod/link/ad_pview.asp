<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../inc/home/func3.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>查看中心</title>
<script src="../../ext/api/play/jsPlayer.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="config.asp"-->
<%

Act = Request("Act")
ID = RequestS("ID","C",48)
Dim AdvID

If Act="View" Then

  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open "SELECT * FROM [WebAdvert] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
  KeyID = rs("KeyID")
  KeyMod = rs("KeyMod")
  InfType = rs("InfType")
  InfName = rs("InfName")
  InfCont = rs("InfCont")
  InfPath = rs("InfPath")
  InfPara = rs("InfPara")
  end if 
  rs.Close()
  SET rs=Nothing 
  
  ParaA = Split(InfPara,"|")
  Para1 = ParaA(0) : Para2 = ParaA(1) : Para3 = ParaA(2) 
  
  sTabs = gContTab(InfCont) 
  Call gCodeStr(sTabs,InfType,InfPara,InfPath)
  sDiv = "<div id='ImgPlayer_"&InfType&"' style='width:"&Para1&"px; height:"&Para2&"px; overflow:hidden; border:1px solid #CCC'>"
  sCode = "<scr"&"ipt src='../../upfile/sys/xadv/pic_"&InfType&".js'></scr"&"ipt>"
   
  Response.Write "<font style='font-size:18px;'>调用代码:</font> "
  Response.Write "<br>"&Server.HTMLEncode(sDiv)
  Response.Write "<br>"&Server.HTMLEncode(sCode)
  Response.Write "<br>"&Server.HTMLEncode("</div>")
  Response.Write vbcrlf&"<br>效果:<hr>"
  Response.Write vbcrlf&sDiv
  Response.Write vbcrlf&sCode&vbcrlf&"</div>"
  
Else

  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open "SELECT * FROM [WebAdvert] WHERE KeyMod='AdRPic'",conn,1,1 
  if NOT rs.eof then 
  Do While Not rs.EOF

  KeyID = rs("KeyID")
  KeyMod = rs("KeyMod")
  InfType = rs("InfType")
  InfName = rs("InfName")
  InfCont = rs("InfCont")
  InfPath = rs("InfPath")
  InfPara = rs("InfPara")
  
  sTabs = gContTab(InfCont) 
  Call gCodeStr(sTabs,InfType,InfPara,InfPath)
  
  rs.MoveNext()
  Loop
  end if 
  rs.Close()
  SET rs=Nothing 
  
End If


Function gContTab(xCont)
  Dim aCont,aUrl,aPic,j,sCont,sPic,sUrl,i
  aCont = Split(xCont&"(^)(^)","(^)")
  aUrl = Split(aCont(0)&"||||||||||","|")
  aPic = Split(aCont(1)&"||||||||||","|")
  aMsg = Split(aCont(2)&"||||||||||","|")
  j = 0 
  sCont = "" :sUrl = "" :sPic = "" :sMsg = ""
  For i=1 To 8
	If aUrl(i-1)<>"" And aPic(i-1)<>"" Then
       j = j + 1
	   sUrl = sUrl&aUrl(i-1)&"|"
       sPic = sPic&aPic(i-1)&"|"
	   sMsg = sMsg&aMsg(i-1)&"|"
	End If
  Next
  sUrl = sUrl&j
  sPic = sPic&j
  sMsg = sMsg&j
  sCont = sUrl&"(^)"&sPic&"(^)"&sMsg
  sCont = Replace(sCont,"(Config_Path)",Config_Path)
  '' // 处理路径
  gContTab = sCont
End Function


Function gCodeStr(xTabs,xType,xPara,xPath) 
  Dim s,i,j,aUrl,aPic,iExt,xID : s="" : xID=xType
  ParaA = Split(xPara,"|")
  Para1 = ParaA(0) : Para2 = ParaA(1) : Para3 = ParaA(2) 
  If Para3=0 Then  Para3=4096
  If Len(Para3)=1 Then Para3=Para3*1000
  If xTabs="0(^)0(^)0" Then
	 s = "<span style='text-align:center; color:red'>广告位招租！</span>"
  Else
	aCont = Split(xTabs&"(^)(^)","(^)")
	aUrl = Split(aCont(0),"|")
	aPic = Split(aCont(1),"|")
	aMsg = Split(aCont(2),"|")
	If uBound(aUrl)=1 Then
	  iExt = lCase(Right(aPic(0),4))
	  If iExt=".swf" Then
	    iPic = "<embed src='"+aPic(0)+"' width='"+Para1+"' height='"+Para2+"' wmode='transparent'></embed>"
	  Else
	    iPic = "<img src='"&aPic(0)&"' width='"&Para1&"' height='"&Para2&"' alt='"+aMsg(0)+"' border='0'>"
	   If Len(aUrl(0))>3 Then 
	    iPic = "<a href='"&aUrl(0)&"' target='_blank'>"&iPic&"</a>"
	   End If
	  End If
	  s = iPic
	  s = Replace(s,"(Config_Path)",Config_Path)
	Else
	  s="" : iPics="" : iUrls="" : iSubj=""  'Response.Write aCont(0)&"<br>"&aCont(1)
	  For i=0 To uBound(aUrl) - 1
	    iPics = iPics&aPic(i)&"|"
		iUrls = iUrls&aUrl(i)&"|"
		iSubj = iSubj&aMsg(i)&"|"
	  Next
	  s = s&vbcrlf&"ImgPlay_"+xID+"_Pics = """+iPics+""";"
	  s = s&vbcrlf&"ImgPlay_"+xID+"_Subj = """+iSubj+""";"
	  s = s&vbcrlf&"ImgPlay_"+xID+"_Urls = """+iUrls+""";"
	  s = s&vbcrlf&"ImgPlay_"+xID+"_Speed = "+Para3+";"
	  s = s&vbcrlf&"ImgPlay_"+xID+"_Now = -1;"
	  Randomize ' 230~470
	  speed = 230 + Int(150*Rnd()) 
	  s = s&vbcrlf&" setTimeout(""ImgPlayMain('"+xID+"',"+Para1+","+Para2+")"","&speed&");"
    End If
  End If
  Call File_AddJS(xID,s)
End Function

Function File_AddJS(xType,xData)
  Dim s : s = xData
  If(Left(s,1)="<") Then ' Html
    s = Replace(s,"\","\\")
    s = Replace(s,"'","\'")
    s = Replace(s,"""","\""")
    a = Split(s,vbcrlf)
    s = ""
    for i=0 To uBound(a)
      s = s&vbcrlf&"document.write("""&a(i)&""");"
    Next 
  End If
  Call File_Add2(Config_Path&"upfile/sys/xadv/pic_"&xType&".js",s,"UTF-8") 
End Function
%>


</body>
</html>
