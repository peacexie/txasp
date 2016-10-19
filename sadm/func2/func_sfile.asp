<%

' Functions for Save Cont to Files

Function get_TmrPara(xSet)
  Dim sql3,OptFilTyp,trPub,tqPub,sCode,tPara(2) 
  trPub = "<input type='hidden' name='KeyCode' id='KeyCode' value='(KeyCode)'>"
  tqPub = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmrPubX888'")
  tPara(0) = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmr"&ModID&"'")
  tPara(1) = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmq"&ModID&"'")
  If inStr(tPara(0),"(KeyCode)")<=0 Then
    tPara(0) = trPub&vbcrlf&tPara(0)
	'Response.Write "<br>1."
  End If
  If inStr(tPara(0)&tPara(1),"InfPara1")<=0 And inStr(tPara(0)&tPara(1),"InfPara2")<=0 Then
    tPara(0) = trPub
	tPara(1) = tqPub
  End If
  If xSet Then
    sCode = Get_FmtID("mdhnsx","")&"-"&Rnd_ID("KEY",6)
	for i=0 to 1
	  tPara(i) = Replace(tPara(i),"(KeyCode)",sCode)
	  tPara(i) = Replace(tPara(i),"(UsrID)","本站:"&Session("UsrID"))
	next
  Else
    tPara(0) = Replace(tPara(0),"(KeyCode)",KeyCode)
	tPara(1) = Replace(tPara(1),"(KeyCode)",KeyCode)
  End If
  If ModID="PicV125" Then
	sql3 = "SELECT TypID,TypName FROM WebType WHERE TypMod='FilTyp' ORDER BY TypID "
	OptFilTyp = Get_rsOpt(conn,sql3,"")
	tPara(0) = rep_TmpTag1(tPara(0),"OptFilTyp",OptFilTyp)
  End If
  'tPara(0) = Replace(tPara(0),":none","")
  get_TmrPara = tPara
End Function

Function get_1Img(xID,xImg)
  Dim a,img
  a = Split(xImg&"^","^")
  If a(1)="" Then
	img = Config_Path&"img/tool/no_pic_160120.jpg"
  Else
	img = Config_Path&"upfile/"&Replace(xID,"-","/")&"/"&a(1)
  End If
  get_1Img = img  
End Function
'Dim vMsg_TmpErrFile : vMsg_TmpErrFile = "(提示：相关资料没有找到，可能是被屏蔽 或者被删除！)"
Function Show_sfImgs(xImgName,xID)
  Dim i : aImgs = Split(xImgName&"^^^^^^^^^^","^")
  For i=1 To 9
    If aImgs(i)="" Then
	  aImgs(i) = Config_Path&"img/tool/no_pic_160120.jpg"
	Else
	  aImgs(i) = Config_Path&"upfile/"&Replace(xID,"-","/")&"/"&aImgs(i)
	End If
  Next
  Show_sfImgs = aImgs
End Function
Function Show_sfRead(xID,xFile)
  If Config_Cont="DB" Then
    If xFile=".rep.htm" Then
      sData = xxxReply
    Else 'org,out
      sData = xxxCont
    End If
  Else 'Save To File
	sData = File_Read(Config_Path&"upfile/"&Replace(xID,"-","/")&xFile,"utf-8")
	If sData="(File Read Error!)" Then
	  If xFile=".rep.htm" Then
        sData = xxxReply
      Else 'org,out
        sData = xxxCont
      End If
	End If
  End If
  Show_sfRead = sData
End Function
Function Show_sfGbook(xID,xFile)
  Dim idPath : idPath = xID
  If Config_Cont="DB" Then
    If xFile=".rep.htm" Then
      Response.Write xxxReply
    Else 'org,out
      Response.Write xxxCont
    End If
  Else 'Save To File
    If Len(idPath)=24 Then 
      idPath = Config_Path&"upfile/"&Replace(idPath,"-","/")&xFile
    End If
    If fil_exist(idPath) Then
      Server.Execute(idPath)
    Else
     If xxxReply&xxxCont<>"" Then
      If xFile=".rep.htm" Then
        Response.Write xxxReply
      Else 'org,out
        Response.Write xxxCont
      End If
	 Else
	  Response.Write vMsg_TmpErrFile
	 End IF
    End If
  End If
End Function
Function Show_sfData(xID,xFile)
  Dim idPath : idPath = xID
  If Config_Cont="DB" Then
    If xxxCont<>"" Then
      Response.Write xxxCont
	Else
	  Response.Write vMsg_TmpErrFile
	End IF
  Else 'Save To File
    If Len(idPath)=24 Then 
      idPath = Config_Path&"upfile/"&Replace(idPath,"-","/")&"/"
    End If
    If fil_exist(idPath&xFile) Then
     If Right(xFile,4)=".htm" Then
	  Server.Execute(idPath&xFile)
	 Else ' for Inner Docs
	  Response.Write Show_sfRead(xID&xFile,"") 
	 End If
    Else
     If xxxCont<>"" Then
      Response.Write xxxCont
	 Else
	  Response.Write vMsg_TmpErrFile
	 End IF
    End If
  End If
End Function
Function Check_sfData(xPath,xFTab,xMsg)
  Dim i,a,flag,idPath : idPath = xPath
  idPath = Config_Path&"upfile/"&Replace(idPath,"-","/")
  flag = true
  a = Split(xFTab,";")
  For i=0 To uBound(a)
  If a(i)<>"" Then
     If Not fil_exist(idPath&a(i)) Then
	   flag = false
	   Exit For
	 End If
  End If
  Next
  If NOT flag AND xMsg<>"" Then
	Response.Write js_Alert(xMsg,"Back","-1")
	Response.Write vbcrlf&"<div style='color:#F00; text-align:center; margin:50px; padding:50px; '>"
	Response.Write vbcrlf&xMsg&"</div>"
    Response.End()
  End If
  Check_sfData = flag
End Function

'ModID -=> TabID
Function rel_ModTab(xMod)
  Dim aTab,aMod,i,flag,sTab : flag=""
  aTab = Split(Config_dbTab,"|")
  aKey = Split(Config_mdKey,"|")
  sTab = "GboSend"
  For i=0 To uBound(aKey)
	If Left(xMod,3)="Mem" Then
	  flag = "OK"
	  rel_ModTab = "GboInfo"
	  Exit Function
	ElseIf Left(xMod,3)=aKey(i) Then
	  flag = "OK"
	  rel_ModTab = aTab(i)
	  Exit Function
	End If
  Next
  If flag="" Then
    'Response.Write vbcrlf&"<br>Error: Data Path!<br>"
  End If
  rel_ModTab = sTab
End Function
'TabID -=> DirID
Function rel_TabPath(xTab)
  Dim aTab,aPath,i,flag : flag=""
  aTab  = Split(Config_dbTab,"|")
  aPath = Split(Config_upDir,"|")
  For i=0 To uBound(aTab)
    If aTab(i)=xTab Then
	  flag = "OK"
	  rel_TabPath = aPath(i)
	  Exit Function
	End If
  Next
  If flag="" Then
    Response.Write vbcrlf&"<br>Error: Data Path!<br>"
  End If
End Function
'KeyID -=> TabID
Function rel_IDTab(xID)
  Dim aTab,aPath,i,sPath,flag : flag=""
  sPath = Split(xID&"-","-")(0)
  aTab  = Split(Config_dbTab,"|")
  aPath = Split(Config_upDir,"|")
  For i=0 To uBound(aTab)
    If aPath(i)=sPath Then
	  flag = "OK"
	  rel_IDTab = aTab(i)
	  Exit Function
	End If
  Next
  If flag="" Then
    Response.Write vbcrlf&"<br>Error: Data Table!<br>"
  End If
End Function


Sub add_sfFile()
  If Config_Cont = "DB" Then
    Exit Sub
  End If
  If ID<>"" And KeyID="" Then
    KeyID = ID
  End If
  'Response.Write upRoot&KeyID&InfCont
  If ModTab="GboSend" Then
    Call fold_add9(upRoot,KeyID,1) 
	Call File_Add2(upPath&".out.htm",InfCont,"utf-8")
  ElseIf Left(ModTab,3)="Gbo" Then
    Call fold_add9(upRoot,KeyID,1) 
	Call File_Add2(upPath&".org.htm",InfCont,"utf-8")
	If InfReply<>"" Then
    Call File_Add2(upPath&".rep.htm",InfReply,"utf-8")
	End If
  ElseIf ModTab="DocsNews" Then
	Call fold_add9(upRoot,KeyID,0)
	Call File_Add2(upPath&"fcont.htx",InfCont,"utf-8")
  Else 'News,Pics,BBS,Trade
	Call fold_add9(upRoot,KeyID,0)
	Call File_Add2(upPath&"fcont.htm",InfCont,"utf-8")
	If inStr("InfoNews;InfoPics;TradeInfo",ModTab) Then
	  Call add_TmpHtml(ModTab,KeyID,InfCont)
	Else 
	  '//////  ' for BBS
	End If
	'Response.Write upRoot&KeyID&InfCont&upPath
  End If
End Sub
' for 'News,Pics,Vdos,Jobs...[公文,论坛,供求]
Function del_sfDir(xTab,sID)
  Dim aID,i
  sID = Replace(Replace(Replace(sID,",",";")," ",""),"'","")
  aID = Split(sID,";")
  For i=0 To uBound(aID)
  If aID(i)<>"" Then
	Call rs_DoSql(conn,"DELETE FROM "&xTab&" WHERE KeyID='"&aID(i)&"' ")
	Call fold_del(upRoot&Replace(aID(i),"-","/")&"/")
	Call rs_DoSql(conn,"DELETE FROM InfoPhoto WHERE KeyRe='"&aID(i)&"' ")
  End If
  Next
End Function
' for 'Gbook,Reply,GboSend
Function del_sfCont(xTab,sID)
  Dim aID,i
  sID = Replace(Replace(Replace(sID,",",";")," ",""),"'","")
  aID = Split(sID,";")
  For i=0 To uBound(aID)
  If aID(i)<>"" Then
	Call rs_DoSql(conn,"DELETE FROM "&xTab&" WHERE KeyID='"&aID(i)&"' ")
	If xTab="GboSend" Then
	  Call fil_del(upRoot&Replace(aID(i),"-","/")&".out.htm")
	Else
	  Call fil_del(upRoot&Replace(aID(i),"-","/")&".org.htm")
	  Call fil_del(upRoot&Replace(aID(i),"-","/")&".rep.htm")
	End If
  End If
  Next
End Function


Sub get_TmpLink()
  Dim i,aImgs,aPara,sGap,sDir
  sGaps = "^^^^^^^^^^"
  aImgs = Split(ImgName&sGaps,            "^")
  aPara = Split(InfPara&sGaps&sGaps&sGaps,"^")
  tmpLink = get_TmpID("Link",KeyMod&";"&InfType)
  'Response.Write tmpLink
  If tmpLink="(Null)" Then
    InfLink = tmpLink
  ElseIf Left(tmpLink,5)="[dir/" Then
    sDir = Config_Path&"upfile/"&Replace(KeyID,"-","/") 
    InfLink = "href='"&Replace(tmpLink,"[dir",sDir)&"'"
  Else
   For i=1 To 9
    aImgs(i) = Config_Path&"upfile/"&Replace(KeyID,"-","/")&"/"&aImgs(i)
	tmpLink = Replace(tmpLink,"(ImgNam"&i&")",aImgs(i))
	'Response.Write "<br>"&i&"--"&aImgs(i)
   Next
   For i=1 To 24
    tmpLink = Replace(tmpLink,"(InfPar"&i&")",aPara(i))
   Next
    tmpLink = Replace(tmpLink,"(KeyID)",KeyID)
    tmpLink = Replace(tmpLink,"(InfType)",InfType)
    tmpLink = Replace(tmpLink,"(KeyMod)",KeyMod)
    InfLink = "href='"&tmpLink&"'"
  End If
End Sub
Function get_TmpID(xTab,xTyps)
  Dim i,j,aItm,aTab
  xTyps = Replace(xTyps,",",";")
  xTyps = Replace(xTyps,";;",";")
  aItm = Split(xTyps,";")
  aTab = Eval("tms"&xTab&"Tab")
  'Eval("aTab = tms"&xTab&"Tab")
  'Response.Write "tms"&xTab&"Tab"
  For i=uBound(aItm) To 0 Step -1
  If aItm(i)<>"" Then
	For j=0 To uBound(aTab)
	  If aTab(j)<>"" Then
	  If inStr(aTab(j),aItm(i)&";")>0 Then
	    p = inStr(aTab(j),":")
		get_TmpID = Left(aTab(j),p-1)
		'Response.Write "tms"&xTab&"Tab ----- "
	    Exit Function
	  End If
	  End If
	Next
  End If
  Next
  'Response.Write "Def]"
  If xTab="Show" Then
    get_TmpID = "1News" '默认 经典:标题+内容(新闻) --- 6UD:
  ElseIf xTab="List" Then
    get_TmpID = "News" '默认 经典:新闻列表
  ElseIf xTab="Link" Then
    get_TmpID = "iview.asp?KeyID=(KeyID)" '默认 简便:?本页
  ElseIf xTab="Rem" Then
    get_TmpID = Mid(tmsRemTab(3),3,1) '"Y" '默认 Remark
  ElseIf xTab="Vote" Then
    get_TmpID = Mid(tmsVoteTab(3),3,1) '"Y" '默认 Vote
  ElseIf xTab="Next" Then
    get_TmpID = Mid(tmsNextTab(4),3,1) '"T" '默认 Next
  Else
    get_TmpID = "---" '默认
  End If
End Function

Function get_TmpMain(xCode)
  sTmp = File_Read(Config_Path&"/pfile/temc/vmain.htm","utf-8")
  p1 = inStr(sTmp,"[!-- Temp["&xCode&"] Start --]") 
  If p1>0 Then
	p2 = inStr(sTmp,"[!-- Temp["&xCode&"] End --]")
	n1 = Len("[!-- Temp["&xCode&"] Start --]</div>")
	n2 = Len("<div style='clear:both'>")
	p1 = p1 + n1
	p2 = p2 - n2
	sTmp = Mid(sTmp,p1,p2-p1)
  Else
    sTmp = ""
  End If
  get_TmpMain = sTmp
End Function
Function get_TmpHtml(xLay)
  Dim xTyps,aItm,sTemp
  xTyps = xLay
  xTyps = Replace(xTyps,";;;",";")
  xTyps = Replace(xTyps,";;",";")
  aItm = Split(xTyps,";")
  sTemp = ""
  For i=uBound(aItm) To 0 Step -1
  If aItm(i)<>"" Then
	If fil_exist(Config_Path&"/pfile/temc/t"&aItm(i)&".htm") Then
	  sTemp = File_Read(Config_Path&"/pfile/temc/t"&aItm(i)&".htm","utf-8")
	  Exit For
	End If
  End If
  Next
  If sTemp="" Then
    sTemp = File_Read(Config_Path&"/pfile/temc/vframe.htm","utf-8")
  End If
  get_TmpHtml = sTemp
End Function

Dim tagTemp1 : tagTemp1 = "<{$"
Dim tagTemp2 : tagTemp2 = "$}>"
Function rep_TmpTags(xStr) 
  Dim n1,n2,n,iTag,iVal',xStr
  'xStr = xData
  n1 = inStr(xStr,tagTemp1)+Len(tagTemp1)
  n2 = inStr(xStr,tagTemp2)
  n = n2-n1
  if n1>0 AND n2>n1 AND n>0 then
    iTag = Mid(xStr,n1,n)
	If iTag<>"" Then
	  iVal = Eval(iTag)
	  If iVal&""="" Then iVal = "{$"&iTag&"$}"
	  xStr = Replace(xStr,tagTemp1&iTag&tagTemp2,iVal)
	  xStr = Replace(xStr,"<"&"%="&iTag&"%"&">",iVal)
    End If
    rep_TmpTags = rep_TmpTags(xStr) 
  end if
  rep_TmpTags = xStr
End Function
Function rep_TmpTag1(xStr,xTag,xVal)
  Dim tmp : xVal=xVal&""
  tmp = Replace(xStr,tagTemp1&xTag&tagTemp2,xVal)
  tmp = Replace(tmp,"<"&"%="&xTag&"%"&">",xVal)
  rep_TmpTag1 = tmp
End Function

Function add_TmpHtml(xTab,xID,xCont)
  Dim rs
  Dim KeyID,KeyMod,KeyCode,InfType,InfTyp2,InfSubj,InfCont,InfPara,aPara
  Dim SetRead,SetSubj,SetHot,SetTop,SetShow,ImgName,LogATime
  Dim tShow,sMain,sView,sItm,sSet,sPub,aItm,i
  If Config_Cont <> "Html" Then
    Exit Function
  End If
  '得到 DB中的字段值
  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open "SELECT * FROM "&xTab&" WHERE KeyID='"&xID&"'",conn,1,1 
    If NOT rs.eof then 
  KeyID = rs("KeyID")
  KeyMod = rs("KeyMod")
  KeyCode = rs("KeyCode")
  InfType = rs("InfType")
  InfTyp2 = rs("InfTyp2")
  InfSubj = Show_Form(rs("InfSubj"))
  InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
  SetRead = rs("SetRead")
  SetSubj = rs("SetSubj")
  SetHot = rs("SetHot")
  SetTop = rs("SetTop")
  SetShow = rs("SetShow")
  ImgName = rs("ImgName")&"" : aImgs = Show_sfImgs(ImgName,KeyID)
  LogATime = rs("LogATime") 
    End If
  rs.Close()
  SET rs=Nothing 
  '得到 内容模版和内容Frame
  tShow = get_TmpID("Show",KeyMod&";"&InfType) 
  sMain = get_TmpMain(tShow) ' 
  sView = get_TmpHtml(KeyMod&";"&InfType) :Response.Write sView
  sView = Replace(sView,"<{$boxMain$}>",sMain)
  sView = Replace(sView,"<{$InfCont$}>",xCont) '"<!"&"--#include"&" file=""fcont.htm""-->"
  '替换 标签
  sItm = "KeyID,KeyMod,KeyCode,InfType,InfTyp2,InfSubj,InfCont,InfPara,"
  sSet = "SetRead,SetSubj,SetHot,SetTop,SetShow,LogATime,"
  sPub = "Config_Path,Config_Name"
  aItm = Split(sItm&sSet&sPub,",")
  For i=0 To uBound(aItm)
   If aItm(i)&""<>"" Then
	If Eval(aItm(i))&""<>"" Then
	  sView = rep_TmpTag1(sView,aItm(i),Eval(aItm(i))) 
	End If
   End If
  Next
  '替换 参数和图片
  For i=1 To 96
    sView = Replace(sView,tagTemp1&"InfPar"&i&tagTemp2,aPara(i))
  Next
  For i=1 To 9
    sView = Replace(sView,tagTemp1&"ImgNam"&i&tagTemp2,aImgs(i))
  Next
  '替换 其它
  sView = rep_TmpTag1(sView,"ModID",KeyMod)
  sView = rep_TmpTag1(sView,"MDName",rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&KeyMod&"'"))
  sView = Replace(sView,tagTemp1&aItm(i)&tagTemp2,Eval(aItm(i)))
  Call fold_add9(upRoot,KeyID,0)
  Call File_Add2(upPath&"index.htm",sView,"utf-8")
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>
