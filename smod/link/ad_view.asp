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
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<link href="../../img/rnd_nid/box_nid.css" rel="stylesheet" type="text/css">
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../../inc/home/jsadv.js" type="text/javascript"></script>
<style type="text/css">
.AdvPClose {
    width:100%;
	font-size:12px;
	line-height:100%;
	padding:2px 1px;
	cursor:pointer;
	text-align:right;
	border-bottom:1px solid #EEE;
	
}
</style>
</head>
<body>
<!--#include file="config.asp"-->
<%

Act = Request("Act")
ID = RequestS("ID","C",48)
TP = RequestS("TP","C",48)
Dim AdvID

If Act="VMod" Then
 ModID = Request("ModID")
 Response.Write "<center><br>"&gModStr(ModID)&"</center>"
 fCont = "../../img/rnd_nid/rbox_nid.htm"
ElseIf Act="View" Then

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
  
  AdvID = "12"&Rnd_ID("0",4)
  eCont = gEffect(InfName,InfCont,InfType,InfPath,InfPara)
  eArr = Split(eCont,"($Span)")
  Response.Write "<div style='width:100%;height:800px;'>&nbsp;</div>"
  Response.Write eArr(0)&eArr(1)
  Response.Write vbcrlf&"<script type='text/javascript'>"&eArr(2)&vbcrlf&"</script>"

Else

  If TP="" Then 
    Response.Write "请选择类别！"
	Response.End()
  End If

  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open "SELECT * FROM [WebAdvert] WHERE KeyMod='Advert' AND InfType LIKE '"&TP&"%'",conn,1,1 
  if NOT rs.eof then 
  iAdv = 11 
  ctAdv = ""
  jsAdv = ""
  Do While Not rs.EOF
  iAdv = iAdv + 1
  KeyID = rs("KeyID")
  KeyMod = rs("KeyMod")
  InfType = rs("InfType")
  InfName = rs("InfName")
  InfCont = rs("InfCont")
  InfPath = rs("InfPath")
  InfPara = rs("InfPara")
  
  AdvID = iAdv&Rnd_ID("0",4)
  eCont = gEffect(InfName,InfCont,InfType,InfPath,InfPara)
  eArr = Split(eCont,"($Span)")
  ctAdv = ctAdv&vbcrlf&vbcrlf& eArr(1)
  jsAdv = jsAdv&vbcrlf&vbcrlf& eArr(2)
  
  rs.MoveNext()
  Loop
  end if 
  rs.Close()
  SET rs=Nothing 
  
  Call File_Add2("../../upfile/sys/xadv/da_"&TP&".asp",ctAdv,"UTF-8")
  Call File_Add2(Config_Path&"upfile/sys/xadv/da_"&TP&".js",jsAdv,"UTF-8")
  
  Response.Write "<div style='width:100%;height:800px;padding:24px;padding-left:120px;'>"
  Response.Write "调用：<br>&lt;!--#include file=&quot;/upfile/sys/xadv/da_"&TP&".asp&quot;--&gt;<br />"
  Response.Write "&lt;script language=&quot;javascript&quot; src=&quot;/upfile/sys/xadv/da_"&TP&".js&quot; type=&quot;text/javascript&quot;&gt;&lt;/script&gt;<br />"  
  Response.Write "</div>"
  Response.Write ctAdv
  Response.Write vbcrlf&"<script type='text/javascript'>"&jsAdv&vbcrlf&"</script>"

End If



Function gEffect(sName,sCont,sType,sPath,sOut)
  Dim InfName,InfType
  
 NameA = Split(sName,"|") ':Response.Write "<br>"&sOut
 ContA = Split(sCont,"|")
 TypeA = Split(sType,"|")
 PathA = Split(sPath,"|")
 OutA = Split(sOut,"|")
 oWidth=OutA(0)    : oHeight=OutA(1)   : oOffX=OutA(2)  : oOffY=OutA(3)  : oRLeft=OutA(4)
 InfType=TypeA(0)  : InfTyp2=TypeA(1)  : InfTyp3=TypeA(2)

 rCont=""
 rJS=""
 If InfTyp2="AdvPair" Then 
  
  rCont=rCont&vbcrlf&"<div id='AdvP01' style='width:"&oWidth&"px;left:"&oOffX&"px;top:"&oOffY&"px;background-color:#FFF;border:1px solid #DDD;position:absolute;'>"
  rCont=rCont&vbcrlf&"<div class='AdvPClose' onclick=""ShowDiv('AdvP01')"">关闭 X&nbsp;</div>"&gContStr(ContA(0),oWidth,oHeight,PathA(0))&"</div>"
  rCont=rCont&vbcrlf&"<div id='AdvP02' style='width:"&oWidth&"px;right:"&oOffX&"px;top:"&oOffY&"px;background-color:#FFF;border:1px solid #DDD;position:absolute;'>"
  rCont=rCont&vbcrlf&"<div class='AdvPClose' onclick=""ShowDiv('AdvP02')"">关闭 X&nbsp;</div>"&gContStr(ContA(1),oWidth,oHeight,PathA(1))&"</div>"
  rJS=rJS&vbcrlf&"function AdvPair1(xDiv){AdvPTopY1=AdvSideY(xDiv,AdvPTopY1);}"
  rJS=rJS&vbcrlf&"AdvPTopY1=0;"
  rJS=rJS&vbcrlf&"window.setInterval(""AdvPair1('AdvP01')"",31);"
  rJS=rJS&vbcrlf&"window.setInterval(""AdvPair1('AdvP02')"",31);"
 
 ElseIf InfTyp2="AdvJJCC" Then
 
  rCont=rCont&vbcrlf&"<div id='AdvP05' style='width:"&oWidth&"px;left:"&oOffX&"px;top:"&oOffY&"px;background-color:#FFF;border:1px solid #DDD;position:absolute;'>"
  rCont=rCont&vbcrlf&"<div class='AdvPClose' onclick=""ShowDiv('AdvP05')"">关闭 X&nbsp;</div>"&gContStr(ContA(0),oWidth,oHeight,PathA(0))&"</div>"
  rCont=rCont&vbcrlf&"<div id='AdvP06' style='width:"&oWidth&"px;right:"&oOffX&"px;top:"&oOffY&"px;background-color:#FFF;border:1px solid #DDD;position:absolute;'>"
  rCont=rCont&vbcrlf&"<div class='AdvPClose' onclick=""ShowDiv('AdvP06')"">关闭 X&nbsp;</div>"&gContStr(ContA(1),oWidth,oHeight,PathA(1))&"</div>"
  rJS=rJS&vbcrlf&"function AdvPair5(xDiv){AdvPTopY5=AdvSideY(xDiv,AdvPTopY5);}"
  rJS=rJS&vbcrlf&"AdvPTopY5=0;"
  rJS=rJS&vbcrlf&"window.setInterval(""AdvPair5('AdvP05')"",37);"
  rJS=rJS&vbcrlf&"window.setInterval(""AdvPair5('AdvP06')"",37);"
  
 ElseIf InfTyp2="Float01" Then

  tCont=gContStr(ContA(0),oWidth,oHeight,PathA(0)) '内容
  rCont=gModStr(TypeA(2)) '摸板
  rCont=Replace(rCont,"($Div)","AdvFloat01")
  rCont=Replace(rCont,"width=""200""","width='"&oWidth&"'")
  rCont=Replace(rCont,"width=""130""","width='"&oWidth&"'")
  rCont=Replace(rCont,"Test标题",NameA(0))
  rCont=Replace(rCont,"内容测试,","")
  rCont=Replace(rCont,"($Date)",NameA(1))
  rCont=Replace(rCont,"($Url)",PathA(0))
  rCont=Replace(rCont,"($Content)",tCont)
  rCont=vbcrlf&"<div id=AdvFloat01 style='width:"&oWidth&"px;left:120px;background-color:#FFF;position:absolute;'>"&rCont&"</div>"
  rJS=rJS&vbcrlf&"var xF01="&oOffX&",  yF01="&oOffY&";"
  rJS=rJS&vbcrlf&"var F01Obj;          F01Obj = getElmID('AdvFloat01');"
  rJS=rJS&vbcrlf&"var IntF01 = setInterval(""Float01()"", F01Delay);"
  rJS=rJS&vbcrlf&"F01Obj.onmouseover=function(){clearInterval(IntF01);} "
  rJS=rJS&vbcrlf&"F01Obj.onmouseout=function(){IntF01=setInterval(""Float01()"", F01Delay);} "
  
 ElseIf InfTyp2="Float02" Then
 
  tCont=gContStr(ContA(0),oWidth,oHeight,PathA(0)) '内容
  rCont=gModStr(TypeA(2)) '摸板
  rCont=Replace(rCont,"($Div)","AdvFloat02")
  rCont=Replace(rCont,"width=""200""","width='"&oWidth&"'")
  rCont=Replace(rCont,"Test标题",NameA(0))
  rCont=Replace(rCont,"内容测试,","")
  rCont=Replace(rCont,"($Date)",NameA(1))
  rCont=Replace(rCont,"($Url)",PathA(0))
  rCont=Replace(rCont,"($Content)",tCont)
  rCont=vbcrlf&"<div id=AdvFloat02 style='width:"&oWidth&"px;left:240px;background-color:#FFF;position:absolute;' onMouseOver='f2Pause=false;' onMouseOut='f2Pause=true;'>"&rCont&"</div>"
  rJS=rJS&vbcrlf&"var f2PosX="&oOffX&", f2PosY="&oOffY&";"
  rJS=rJS&vbcrlf&"f2Int = setInterval('Float02()', f2Delay); "
  rJS=rJS&vbcrlf&"var f2Yon=0,          f2Xon=0; "
  rJS=rJS&vbcrlf&"var f2Int,            f2Pause = true; "

 ElseIf InfTyp2="AdvLRXX" Then
 
  tCont=gContStr(ContA(0),oWidth,oHeight,PathA(0)) '内容 InfTyp3=Left/Right
  rCont=vbcrlf&"<div id='AdvSDiv001' style='width:"&oWidth&"px;"&oRLeft&":"&oOffX&"px;top:"&oOffY&"px;background-color:#FFF;border:1px solid #EEE;position:absolute;'>"
  rCont=rCont&"<div class='AdvPClose' onclick=""ShowDiv('AdvSDiv001')"">[X]</div>($Cont)</div>" '
  rJS=rJS&vbcrlf&"function AdvSide001(xDiv){AdvSY001=AdvSideY(xDiv,AdvSY001);}"
  rJS=rJS&vbcrlf&"AdvSY001=0;window.setInterval(""AdvSide001('AdvSDiv001')"",41);"
  
  rJS = Replace(rJS,"001","0"&AdvID)
  rCont = Replace(rCont,"001","0"&AdvID)
  rCont = Replace(rCont,"($Cont)",tCont)
  rCont=Replace(rCont,"($Div)",Replace("AdvSDiv001","001","0"&AdvID))

 ElseIf InfTyp2="AdvLRQQ" Then
 
  rCont=gModStr(TypeA(2)) 
  aQQ = Split(ContA(0),":") : sQQ=""
  For i=0 To uBound(aQQ)
	sQQ = sQQ&gItmStr(aQQ(i))
  Next
  rCont = Replace(rCont,"<!--ItemList-->",sQQ)
  rCont = Replace(rCont,"width=""130""","width='"&oWidth&"'")
  rCont = Replace(rCont,"Test标题",NameA(0))
  tCont = rCont
  rCont = vbcrlf&"<div id='AdvQDiv001' style='width:"&oWidth&"px;"&oRLeft&":"&oOffX&"px;top:"&oOffY&"px;background-color:#FFF;position:absolute;'>($Cont)</div>"
  rJS=rJS&vbcrlf&"function AdvQF001(xDiv){AdvQQ001=AdvSideY(xDiv,AdvQQ001);}"
  rJS=rJS&vbcrlf&"AdvQQ001=0;window.setInterval(""AdvQF001('AdvQDiv001')"",43);"
  
  rJS = Replace(rJS,"001","0"&AdvID)
  rCont = Replace(rCont,"001","0"&AdvID)
  rCont = Replace(rCont,"($Cont)",tCont)
  rCont=Replace(rCont,"($Div)",Replace("AdvQDiv001","001","0"&AdvID))

 End If
  'rCont=Replace(rCont,"($Div)",AdvID)
  gEffect = TypeA(0)&"($Span)"&rCont&"($Span)"&rJS 
End Function

Function gContStr(xCont,xW,xH,xUrl)
  'Response.Write VBCRLF&"A0"&xCont
  Dim s : s=Trim(xCont&"")
  If Right(s,1)=">" And ( Left(lCase(s),4)="<img" Or Left(lCase(s),6)="<embed" ) Then
      s="<a href='"&xUrl&"' target='_blank'>"&Show_Text(s)&"</a>"
  ''//  Left(s,7)="http://" And
  ElseIf ( Right(lCase(s),4)=".swf" Or Right(lCase(s),4)=".jpg" Or Right(lCase(s),4)=".gif" ) Then
	If Right(lCase(s),4)=".swf" Then
	  s="<embed src='"&xCont&"' quality='high' type='application/x-shockwave-flash' width="&xW&" xheight="&xH&" wmode='transparent'></embed>"
	Else
	  s="<a href='"&xUrl&"' target='_blank'><img src='"&xCont&"' width="&xW&" xheight="&xH&" border='0'/></a>"
	End If 
  Else
      s="<a href='"&xUrl&"' target='_blank'>"&Show_Text(s)&"</a>"
  End If
  s = Replace(s,"(Config_Path)",Config_Path)
  gContStr=s
End Function


Function gModStr(xMD)
  gModStr = ListGTemp("Code",xMD,"/img/rnd_nid/rbox_nid.htm")
End Function

'' 80893510:Msn+peace@msn.com:Skype+peace0769:Ali+peace:Tao+peaceTaobao
Function gItmStr(xID)
Dim p,CD,t : CD = "QQ"
  If inStr(xID,"+")>0 Then
    p = inStr(xID,"+")
	CD = Left(xID,p-1)
  End If
  t = ListGTemp("Para","tmMsg"&CD,"")
  t = Replace(t,"/img/",Replace(Config_Path&"/img/","//","/"))
  ID = Replace(xID,CD&"+","")
  t = Replace(t,"("&CD&"ID)",ID)
  t = "<div style='padding:2px 0px 0px 0px'>"&t&"</div>"
  gItmStr = t
End Function

%>

</body>
</html>
