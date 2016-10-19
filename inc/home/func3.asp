<%

Function get_1Pic(xMod,xType,x1,x2,xUrl,xWhr)
  Dim sql,sID,sSubj,sSBak,sImg
  'If xWhr="" Then xWhr=" AND SetHot='Y' "
  sql = "SELECT TOP 1 * FROM "&rel_ModTab(xMod)&""
  sql =sql& " WHERE KeyMod='"&xMod&"' AND SetShow='Y' AND InfType LIKE '"&xType&"%' AND LEN(ImgName)>15 "&xWhr&" "
  sql =sql& " ORDER BY LogATime DESC " 
  rs.Open sql,conn,1,1 
  If NOT rs.EOF Then
   sID = rs("KeyID")
   sSubj = rs("InfSubj") :sSBak = Show_Form(sSubj)
   sSubj = Show_SLen(sSubj,120)
   sSubj = Show_sTitle(sSubj,rs("SetSubj"))
   sImg  = get_1Img(sID,rs("ImgName")) 
   If inStr(sImg,"/img/tool/no_pic_")>0 Then
     sImg = "<img src='"&sImg&"' width='"&x1&"' height='"&x2&"' alt='无图片' />"
   Else
     If xUrl="" Then xUrl="iview.asp?KeyID=(KeyID)"
	 xUrl = Replace(xUrl,"(KeyID)",sID)
	 sImg = "<a href='"&xUrl&"'><img src='"&sImg&"' width='"&x1&"' height='"&x2&"' alt='"&sSBak&"' onload='javascript:setImgSize(this);' /></a>"
   End If
  Else
     sImg = "<img src='"&get_1Img("","") &"' width='"&x1&"' height='"&x2&"' alt='无图片' />"
  End If
  If inStr(sImg,"/img/tool/no_pic_")>0 Then
    If int(x2)>int(x1) Then
	  xUrl = Replace(xUrl,"/no_pic_160120.jpg","/no_pic_120150.jpg")
	End If
  End If
  rs.Close()
  get_1Pic = sImg
End Function

' 得到切换图参数：playpic.swf/playpid.swf播放;
Function get_Player(xSql,xLnk)
 Dim a(3),rp1Img,rp1Url,rp1Txt,sID,sSubj,sImg,iLink
 rs.Open xSql,conn,1,1 
 rp1Img="":rp1Url="":rp1Txt=""
 Do While NOT rs.EOF
  sID = rs("KeyID") : iLink = Replace(xLnk,"($KeyID)",sID)
  sSubj = Show_SLen(rs("InfSubj"),120)
  sSubj = Replace(sSubj,"'","")
  sImg  = get_1Img(sID,rs("ImgName")) 
  rp1Img = rp1Img&sImg&"|"
  rp1Url = rp1Url&iLink&"|"
  rp1Txt = rp1Txt&sSubj&"|"
 rs.MoveNext
 Loop
 rs.Close()
 If rp1Url<>"" Then
   a(0) = Replace(rp1Img&"$","|$","")
   a(1) = Replace(rp1Url&"$","|$","")
   a(2) = Replace(rp1Txt&"$","|$","")
 Else
   a(0)="" : a(1)="" : a(2)=""
 End If
 get_Player = a
End Function 

' 得到切换图参数：jsPlay02.js播放;
Function get_Play02(xSql,xLnk)
 Dim s,sID,sSubj,sImg,iLink : s=""
 rs.Open sql,conn,1,1 
 Do While NOT rs.EOF
  sID = rs("KeyID") : iLink = Replace(xLnk,"($KeyID)",sID)
  sSubj = Show_SLen(rs("InfSubj"),120)
  sSubj = Replace(Replace(sSubj,"'",""),"""","")
  sImg  = get_1Img(sID,rs("ImgName")) 
  s = s&vbcrlf&"jsPlay02.addItem('"&sSubj&"','"&iLink&"','"&sImg&"');"
 rs.MoveNext
 Loop
 rs.Close()
 get_Play02 = s
End Function 

' 得到类别列表的一部分 含 xTM;
Function GetItemPart(xMD,xTM)            
 Dim i,j,aCode,aName,aNam2,aLay,sCode,sName,sNam2,sLay
 aCode = Split(Eval("s"&xMD&"Code"),"|")
 aName = Split(Eval("s"&xMD&"Name"),"|")
 aLay = Split(Eval("s"&xMD&"Lay"),"|")
 aNam2 = Split(Eval("s"&xMD&"Nam2"),"|")
 j=0
 'Response.Write xMD&xTM&Eval("s"&xMD&"Code")&"--"
 for i = 0 to uBound(aCode)-1	
   If inStr(aLay(i),xTM&";")>0 Then 
     j=j+1
   End If
   If inStr(aLay(i),xTM&";")>0 And j>1 Then 
	 sCode = sCode&aCode(i)&"|"
	 sName = sName&aName(i)&"|"
	 sLay = sLay&Replace(aLay(i),xTM&";","")&"|"
	 sNam2 = sNam2&aNam2(i)&"|"
   End If
 next
 'Response.Write sCode&"^"&sName&"^"&sLay&"^"&sNam2
 GetItemPart=sCode&"^"&sName&"^"&sLay&"^"&sNam2
End Function

' 得到 折叠菜单
Function GetItemTree(xMD,xUrl,xTM)            
 Dim i,s,aCode,aName,aLay,TypID,TypName,TypLayer,nLayer,nLayStr,NextLay,NextDeep,j,sTmp
 If xTM="" Then
   aCode = Split(Eval("s"&xMD&"Code"),"|")
   aName = Split(Eval("s"&xMD&"Name"),"|")
   aLay = Split(Eval("s"&xMD&"Lay"),"|")
 Else
   sTmp=GetItemPart(xMD,xTM)
   aTmp=Split(sTmp,"^")
   aCode = Split(aTmp(0),"|")
   aName = Split(aTmp(1),"|")
   aLay = Split(aTmp(2),"|")
 End If
 for i = 0 to uBound(aCode)-1	
	TypID = aCode(i)
	TypName = aName(i)
	TypLayer = aLay(i)
		nLayer = Len(TypLayer)-Len(Replace(TypLayer,";",""))
		nLayStr = ""
		NextLay = aLay(i+1)
		NextDeep = Len(NextLay)-Len(Replace(NextLay,";",""))
		If nLayer > 0 Then
		  For j = 1 To nLayer-1
		  If j=1 Then '//////////////////////////////////////////////// 
		   If Int(nLayer)=2 Then
		    If NextLay="" Then
			 nLayStr = nLayStr & "<img src='../img/tree/tree_c.gif' width='17' height='16'>"
			Else
			 nLayStr = nLayStr & "<img src='../img/tree/tree_b.gif' width='17' height='16'>"
			End If
		   Else
			 nLayStr = nLayStr & "<img src='../img/tree/tree_a.gif' width='17' height='16'>"
		   End If
		  ElseIf j=nLayer-1 Then
		   If NextLay="" Then
		     nLayStr = nLayStr & "<img src='../img/tree/tree_c.gif' width='17' height='16'>"
		   ElseIf Int(NextDeep)<>Int(nLayer) Then
		     nLayStr = nLayStr & "<img src='../img/tree/tree_c.gif' width='17' height='16'>"
		   Else
		     nLayStr = nLayStr & "<img src='../img/tree/tree_b.gif' width='17' height='16'>"
		   End If
		  Else
		     nLayStr = nLayStr & "<img src='../img/blank.gif' width='17' height='16'>"
		  End If
		  Next
		End If
		If NextDeep>nLayer Then 'If rs_Count(conn,"WebTyps WHERE TypMod='"&MD&"' AND TypLayer LIKE '"&TypLayer&"%'")>1 Then
		  nLayStr = nLayStr & "<img src='../img/tree/tree_+.gif' width='15' height='15' style='cursor:hand;' onClick=""ChkTree('"&TypLayer&"')"">"
		  sAllLays = sAllLays&TypID&";" ' 展开另一个+时,折叠上一个+
		Else
		  nLayStr = nLayStr & "<img src='../img/tree/tree_-.gif' width='15' height='15'>"
		End If
	  s = s &vbcrlf&nLayStr& "<a href='"&xUrl&"TypID="&TypID&"&TPLay="&TypLayer&"'>"&TypName&"</a> "&xxxLayer&" <br> "
	 If inStr(nLayStr,"tree_+.gif")>0 Then
	  s = s &vbcrlf&"<div id='PeaceM"&TypID&"' style='display:none'>"
	 End if
	 If Int(NextDeep)<Int(nLayer) Then
	  For j = 1 To nLayer-NextDeep
	   s = s &vbcrlf& "</div>"
	  Next	  
	 End if
 next
 GetItemTree = ""&s&""
 Response.Write "<script type='text/javascript'>sLays='"&sAllLays&"';</script> " 
End Function

' 得到 分级 类别列表（多级）
Function GetItemLays(xMD,xUrl,xTM)
Dim i,iDeep,jDeep,iPrev,s,aCode,aName,aLay,sTmp
 If xTM="" Then
   aCode = Split(Eval("s"&xMD&"Code"),"|")
   aName = Split(Eval("s"&xMD&"Name"),"|")
   aLay = Split(Eval("s"&xMD&"Lay"),"|")
 Else
   sTmp=GetItemPart(xMD,xTM)
   aTmp=Split(sTmp,"^")
   aCode = Split(aTmp(0),"|")
   aName = Split(aTmp(1),"|")
   aLay = Split(aTmp(2),"|")
 End If
 for i = 0 to uBound(aCode)-1
  iDeep = Len(aLay(i))-Len(Replace(aLay(i),";",""))
  s=s&vbcrlf&"<div class='SysI0"&iDeep&"'><a href='"&xUrl&"TypID="&aCode(i)&"'>"&aName(i)&"</a></div>"
 next
GetItemLays = s
End Function

' 得到 类别列表（含Flag标记:Sin,Mul,Page,List,Pics,Link,LInn）
Function GetItemList(xMD,xUrl,xTM)
Dim i,s,aCode,aName,iFlag,iUrl
 If xTM="" Then
   aCode = Split(Eval("s"&xMD&"Code"),"|")
   aName = Split(Eval("s"&xMD&"Name"),"|")
   aNam2 = Split(Eval("s"&xMD&"Nam2"),"|")
 Else
   sTmp=GetItemPart(xMD,xTM)
   aTmp=Split(sTmp,"^")
   aCode = Split(aTmp(0),"|")
   aName = Split(aTmp(1),"|")
   aNam2 = Split(aTmp(3),"|")
   'Response.Write sPicS124Code&"--"
 End If
  For i=0 to uBound(aCode)
    If aCode(i)<>"" Then
      iFlag = aNam2(i)
	  If Mid(iFlag,5,1)=")" Then
	   iUrl = Mid(iFlag,6)
	   If Left(iFlag,4)="Link" Then ' Link
		 s=s&vbcrlf&"<div class='SysI01 SysI00'><a href='"&iUrl&"' target='_blank'>"&aName(i)&"</a></div>"
	   Else
	     s=s&vbcrlf&"<div class='SysI01 SysI00'><a href='"&iUrl&"'>"&aName(i)&"</a></div>"
	   End If
	  Else
	     s=s&vbcrlf&"<div class='SysI01 SysI00'><a href='"&xUrl&"TypID="&aCode(i)&"&Flag="&iFlag&"'>"&aName(i)&"</a></div>"
	  End If
    End If
  Next
GetItemList = s
End Function

' 得到 模块名称
Function GetMName(xMD) 'ModID:名称
Dim aCode,aName,i 
 aCode = Split(sysModCode,"|")
 aName = Split(sysModName,"|")
 For i=0 To uBound(aCode)
  If xMD=aCode(i) Then
   GetMName=aName(i)
   Exit For
  End If
 Next
End Function

' 得到 类别名称 或 Lay字串
Function GetTName(xMD,xID,xFlag) 'TypID:名称,S110048;S120112;
Dim aCode,aName,i 
 If inStr(xID,";") Then
   aCode = Split(Eval("s"&xMD&"Lay"),"|")
 Else
   aCode = Split(Eval("s"&xMD&"Code"),"|")
 End If
 If xFlag="Lay" Then
   aName = Split(Eval("s"&xMD&"Lay"),"|")
 Else
   aName = Split(Eval("s"&xMD&"Name"),"|")
 End If
 For i=0 To uBound(aCode)
  If xID=aCode(i) Then
   GetTName=aName(i)
   Exit For
  End If
 Next
End Function

' 得到 类别名称 含级别：衣(服饰)系列 >> 童装 >> 
Function GetNLay(xMD,xID,xLink,xStr1,xStr2) 
Dim aCode,aName,i,tLay,iLink
 If xLink="" Then xLink="info.asp?ModID="&xMD&"&TypID=(TypID)"
 If xStr1="" Then xStr1=" &gt;&gt; "
 aLay = Split(GetTName(xMD,xID,"Lay"),";")
 tLay = ""
 If uBound(aLay)=0 Then 
   tLay = GetTName(xMD,aLay(0),"")
 Else
  For i=0 To uBound(aLay)-1
   iLink = Replace(xLink,"(TypID)",aLay(i))
   iLink = "<a href='"&iLink&"'>"&GetTName(xMD,aLay(i),"")&"</a>"
   tLay = tLay&xStr1&iLink&xStr2
  Next
 End If
 GetNLay=tLay
End Function

' 得到 (图片)连接/////////////////////
Function ListLink(xType,xTemp)
 Dim s,s0,tSql,InfSubj,InfUrl,ImgName,SetSubj : s="" 
 If xTemp="" Then xTemp=rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmp"&xType&"'")
 If xTemp="" Then xTemp="<span class='ItmLPic'><a href='($InfUrl)' target='_blank'>($InfSubj)</a></span>"
 tSql = " SELECT InfSubj,InfUrl,ImgName,SetSubj FROM GboLink WHERE InfType='"&xType&"' ORDER BY SetTop,KeyID DESC "
 rs.Open tSql,conn,1,1 
 Do While NOT rs.EOF
   InfSubj = rs("InfSubj") : InfSBak = Show_Form(InfSubj)
   InfUrl = rs("InfUrl")
   ImgName = rs("ImgName")&""
   SetSubj = rs("SetSubj")
   InfSubj = Show_sTitle(InfSubj,SetSubj)
   s0 = xTemp
   s0 = Replace(s0,"($InfUrl)",InfUrl)
   s0 = Replace(s0,"($InfSubj)",InfSubj)
   s0 = Replace(s0,"($InfSBak)",InfSBak)
   s0 = Replace(s0,"($ImgName)",ImgName)
   s=s&s0
 rs.MoveNext
 Loop
 rs.Close()
 ListLink=s
End Function

' 得到 上下页
Function ListPNext(xTab,xMod,xTyp,xTim,xWhr)
  Dim sql,iID,iSubj,iStr
  sql="SELECT TOP 1 KeyID,InfSubj FROM ["&xTab&"] WHERE "
  If xTyp<>"" Then
    sql=sql& " InfType LIKE '%"&xTyp&"%' "
  Else
    sql=sql& " KeyMod='"&xMod&"' "
  End If
  sql=sql& " AND SetShow='Y' "
  If xWhr=">" Then
    sql=sql& " AND LogATime>"&cfgTimeC&""&xTim&""&cfgTimeC&" ORDER BY LogATime ASC "
  Else
    sql=sql& " AND LogATime<"&cfgTimeC&""&xTim&""&cfgTimeC&" ORDER BY LogATime DESC "
  End If
  rs.Open sql,conn,1,1 
  if NOT rs.EOF then
    iID = rs("KeyID")
	iSubj = Show_Text(rs("InfSubj"))
	iStr = "<a href='?KeyID="&iID&"' target='_self' >"&iSubj&"</a>"
  else
    If vMsg_InfNull&""="" Then vMsg_InfNull="(没有了)"
	iStr = "<font color=gray>"&vMsg_InfNull&"</font>"
  end if 
  rs.Close()
  ListPNext = iStr
End Function

' 得到 摸版：xType:Code,File,Para; 
Function ListGTemp(xType,xID,xPath2)
  Dim s,p1,p2,xPath :xPath=xPath2
  If xType="Para" Then
    ListGTemp=rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='"&xID&"'")
  ElseIf xType="File" Then
    ListGTemp=File_Read(xPath,"utf-8")
  ElseIf xType="Code" Then
   If xPath="" Then 
     xPath = "/img/rnd_nid/rbox_nid.htm"
   End If
     xPath = Replace(Config_Path&xPath,"//","/")
	s=File_Read(xPath,"utf-8") 'File_Read("../../img/rnd_nid/rbox_nid.htm","utf-8")
    s=Replace(s,"../img/",Config_Path&"img/")
    p1=inStr(s,xID&" Start")
    p2=inStr(s,xID&" End")
    p1=p1+Len(xID)+6
    s=Mid(s,p1,p2-p1)
    ListGTemp = s
  Else
    ListGTemp=rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='"&xID&"'")
  End If
End Function

' 得到 摸版列表，用于测试
Function ListTemp(xTemp,xLine)
  Dim i,s : s="" 
  For i=1 To xLine
    s=s&xTemp
  Next
  s = Replace(s,"href=","x=")
  s = Replace(s,"($InfSubj)","<span class=cGray>(无资料,请添加...)</span>")
  If inStr(s,"$LogATime")>0 Then
    s = Replace(s,"$LogATime","<span class=cGray>1900-12-31</span>")
  End If
  ListTemp = s
End Function

' 得到 资料列表，新闻，图片
Function ListPub(xTemp,xLen,xSql)
 Dim s,s0 : s="" 
 Dim KeyID,InfSubj,LogATime,ImgName
 rs.Open xSql,conn,1,1 
 Do While NOT rs.EOF
   KeyID = rs("KeyID")
   InfSubj = Show_SLen(rs("InfSubj"),xLen)
   SetSubj = rs("SetSubj")
   InfSubj = Show_sTitle(InfSubj,SetSubj)
   LogATime = FormatDateTime(rs("LogATime"),2)
   s0 = xTemp
   If inStr(xTemp,"($ImgName)")>0 Then 
     ImgName = get_1Img(KeyID,rs("ImgName"))
	 s0 = Replace(s0,"($ImgName)",Replace(Config_Path&ImgName,"//","/"))
   End If
   s0 = Replace(s0,"($KeyID)",KeyID)
   s0 = Replace(s0,"($InfSubj)",InfSubj)
   s0 = Replace(s0,"($LogATime)",LogATime)
   s=s&s0
 rs.MoveNext
 Loop
 rs.Close()
 If s="" Then s=ListTemp(xTemp,3)
 ListPub=s
End Function


%>