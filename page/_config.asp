<%
'inc file : ../sadm/func2/cch_Class.asp
'Call SubDirect()
Set rs=Server.CreateObject("Adodb.Recordset")
verNow = "2"
'Call sf_Guard()
'Call cchCheck("") 'cchRAM() 'Call cchFile()
%>
<!--#include file="../sadm/func1/func1.asp"-->
<!--#include file="../sadm/func2/func2.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="../sadm/func1/func_file.asp"-->

<!--#include file="../inc/home/func3.asp"-->
<!--#include file="../pfile/lang/vpage.asp"-->

<!--#include file="../upfile/sys/para/tmstyp2.asp" -->
<!--#include file="../upfile/sys/para/tmsPara.asp" -->
<!--#include file="../sadm/func2/func_sfile.asp" -->

<!--#include file="../upfile/sys/config/MTypList.asp"-->
<!--#include file="../upfile/sys/config/_depart.asp"-->
<!--#include file="../upfile/sys/para/seopara.asp"-->

<%

'sysCorp  = rs_Val("","SELECT InfCont FROM InfoNews WHERE InfType like '%H110012;%'")	
'sysPrize = rs_Val("","SELECT InfCont FROM InfoNews WHERE InfType like '%H110020;%'")	
sysFoot  = rs_Val("","SELECT InfCont FROM InfoNews WHERE InfType like '%A110036;%'")
'sysFoot  = filFlagP(sysFoot)

Function filFlagP(xStr)
  Dim xxxCont : xxxCont = xStr
  xxxCont = Replace(xxxCont,"<p>","")
  xxxCont = Replace(xxxCont,"</p>","")
  xxxCont = Replace(xxxCont,"<p>","")
  xxxCont = Replace(xxxCont,"</p>","")
  filFlagP = xxxCont
End Function 

' 得到 第一个 类别ID 'aTP = Split(Eval("s"&MD&"Code"),"|") : TP = aTP(0)
Function GetFirstTypeID(xMD,xTM,xDeep)
Dim i,iDeep,jDeep,iPrev,aCode,aName,aLay,sTmp
 If xTM="" Then
   aCode = Split(Eval("s"&xMD&"Code"),"|")
   aLay = Split(Eval("s"&xMD&"Lay"),"|")
 Else
   sTmp=GetItemPart(xMD,xTM)
   aTmp=Split(sTmp,"^")
   aCode = Split(aTmp(0),"|")
   aLay = Split(aTmp(2),"|")
 End If
 If xDeep="" Then xDeep=1
 for i = 0 to uBound(aCode)-1
   iDeep = Len(aLay(i))-Len(Replace(aLay(i),";",""))
   If(iDeep=Int(xDeep)) Then
     GetFirstTypeID = aCode(i)
	 Exit Function
   End If 
 next
 GetFirstTypeID = ""
End Function

' 得到 子类别列表
smTmp_Spry0 = "<li><a href='/page/info.asp?ModID=(iMod)&TypID=(iTyp)'>(Name)</a></li>"
smTmp_Side1 = "<tr><td height='33' class='pgSub01'><a href='/page/info.asp?ModID=(iMod)&TypID=(iTyp)' class='as06'>(Name)</a></td></tr>"
Function GetSubMenuList(xMD,xUrl)
  Dim i,iDeep,s,aCode,aName,aLay,sTmp,iTmp
  aCode = Split(Eval("s"&xMD&"Code"),"|")
  aName = Split(Eval("s"&xMD&"Name"),"|")
  aLay = Split(Eval("s"&xMD&"Lay"),"|")
  If xUrl="" Then '默认为 SpryMenuBar
    sTmp = smTmp_Spry0
  ElseIf Len(xUrl)>12 Then '自定义
    sTmp = xUrl
  Else
    sTmp = Eval("smTmp_"&xUrl) '如Side1,Spry0
  End If
  for i = 0 to uBound(aCode)-1
    iTmp = Replace(sTmp,"(Name)",aName(i))
    iTmp = Replace(iTmp,"(iMod)",xMD)
    iTmp = Replace(iTmp,"(iTyp)",aCode(i))
    s=s&vbcrlf&iTmp
  next
  If xUrl="" Then s="<ul>"&s&"</ul>"
  GetSubMenuList = s
End Function
Function GetSubMenu01(xMod)
 Dim i,j,aCode,aName,aNam2,aLay,sCode,sName,sNam2,sLay
 aCode = Split(Eval("s"&xMod&"Code"),"|")
 aName = Split(Eval("s"&xMod&"Name"),"|")
 s="" : j=0
 for i = 0 to uBound(aCode)-1	
   s = s&vbcrlf& "<h3><a href='/page/info.asp?ModID="&xMod&"&TypID="&aCode(i)&"'>"&aName(i)&"</a></h3>"
 next
 GetSubMenu01 = s
End Function 
Function GetSubOpt01(xMod,xType)
 Dim i,j,aCode,aName,aNam2,aLay,sCode,sName,sNam2,sLay
 aCode = Split(Eval("s"&xMod&"Code"),"|")
 aName = Split(Eval("s"&xMod&"Name"),"|")
 s="" : j=0
 for i = 0 to uBound(aCode)-1	
     fSel = ""
   If xType=aCode(i) Then
     fSel = " selected "
   End If
   s = s&vbcrlf& "<option value='"&aCode(i)&"' "&fSel&">"&aName(i)&"</option>"
 next
 GetSubOpt01 = s
End Function 

' 得到 折叠菜单:固定2级
Function GetTypeLay2(xMD,xUrl)            
Dim i,iDeep,jDeep,iPrev,s,aCode,aName,aLay,sTmp,sAllLays
   aCode = Split(Eval("s"&xMD&"Code"),"|")
   aName = Split(Eval("s"&xMD&"Name"),"|")
   aLay = Split(Eval("s"&xMD&"Lay"),"|")
 for i = 0 to uBound(aCode)-1
  iDeep = Len(aLay(i))-Len(Replace(aLay(i),";",""))
  jDeep = Len(aLay(i+1))-Len(Replace(aLay(i+1),";",""))
  '  onClick=""ChkTree('"&aCode(i)&"')"" style='display:none;cursor:hand;'
  If iDeep=1 Then
    s=s&vbcrlf&"<div class='SysI0"&iDeep&"' style='cursor:hand;' onClick=""ChkTree('"&aCode(i)&"')""> + "&aName(i)&"</div>"
	s=s&vbcrlf&"[div id='PeaceM"&aCode(i)&"' style='display:none;'>" '
	sAllLays = sAllLays&aCode(i)&";" ' 展开另一个+时,折叠上一个+
  Else
    s=s&vbcrlf&"<div class='SysI0"&iDeep&"'><a href='/page/depart.asp?DepID="&aCode(i)&"&TPLay="&aLay(i)&"'>"&aName(i)&"</a></div>"
  End If
  If jDeep=1 Then
     s=s&vbcrlf&"[/div]"
  Else
     If i=uBound(aCode)-1 Then s=s&vbcrlf&"[/div]"
  End If
 next
 s = Replace(s,"[div ","<div ")
 s = Replace(s,"[/div]","</div>")
 Response.Write "<script type='text/javascript'>sLays='"&sAllLays&"';</script> " 
 GetTypeLay2 = ""&s&""
End Function

%>