<%

Function Get_Title(xHtml)
  Dim px1,px2,sTitle,xStr
  xStr = xHtml
  px1=OutSPos(xStr,"<title>",7) 
  px2=OutSPos(xStr,"</title>",0)
  If px1>0 And px2-px1>0 Then
    sTitle = Mid(xStr,px1,px2-px1)
  End If
  Get_Title = sTitle
End Function

Function Get_Rep(xStr,xTab1,xTab2)
  Dim aRep1,aRep2,sObj,i
  aRep1 = Split(xTab1,"|")
  aRep2 = Split(xTab2,"|")
  For i=0 To uBound(aRep1)
   If aRep1(i)<>"" Then 
    sObj = ""
	If xTab2<>"" Then 
	sObj = aRep2(i)
	End If
	xStr=Replace(xStr,aRep1(i),sObj)
   End If 
  Next
  Get_Rep = xStr 
End Function

Function Get_1Url(xStr,xFlag)
  Dim a,s
  a = Split(xStr,xFlag)
  If uBound(a)>0 Then
    s = a(1)
	b = Split(s," ")
	If uBound(b)>0 Then
	  s = b(0)
	End If
	If Left(s,1)="'" Or Left(s,1)="""" Then
	  s = Replace(s,"""","'")
	  a = Split(s,"'")
	  s = a(1)
	End If
  Else
    s = ""
  End If
  Get_1Url = s
End Function

Function Get_HLinks(xStr,xPat)
  Dim regEx, iMatch, Matches,rStr ' 建立变量。
  rStr = ""
  Set regEx = New RegExp          ' 建立正则表达式。
  regEx.IgnoreCase = True         ' 设置是否区分字符大小写。
  regEx.Global = True             ' 设置全局可用性。 <img name="s" ... />
  regEx.Pattern = xPat '"<a[^>]*(href=[^>]*)[^>]*>([^<]*|<img[^<]*>)</a>" ' 设置模式。
  Set Matches = regEx.Execute(xStr)       ' 执行搜索。
  For Each iMatch In Matches              ' 遍历匹配集合。
    rStr = rStr & iMatch.Value & "||" 
  Next
  'Response.Write rStr
  'regEx.Pattern = "<a[^<]*href=['|""]([^<]*)['|""][^<]*>([^<]*)</a>"
  'rStr = regEx.replace(rStr,"<a href='$1' target='_blank'>$2</a>")
  Set regEx = nothing
  Get_HLinks = rStr
End Function

Function OutSFlag(s,f1,f2)
  Dim s0,p1,p2
  s0 = s
  p1 = OutSPos(s0,f1,Len(f1)) 
  p2 = OutSPos(s0,f2,0)
  If p1>0 And p2>p1 Then
    s0 = Mid(s0,p1,p2-p1)
  Else
    s0 = ""
  End If
  OutSFlag = s0
End Function

Function OutSPos(xCont,xStr,xOffset)
  Dim vCont,vStr,p
  vCont=uCase(xCont)
  vStr=uCase(xStr)
  'p = inStr(vCont,vStr)
  'If p>0 Then p=p+Len(vStr)
  OutSPos=inStr(vCont,vStr)+xOffset
End Function

Function OutPage(Path,xCharSet)
Dim tData,tLen
  tData = OutBody(Path) 
  tLen = Len(tData) ':Response.Write tLen
  if tLen =< 720 then 
    OutPage = OutToStr(tData,xCharSet) 'tData
  else
    OutPage = OutToStr(tData,xCharSet)
  end if 
End function

Function OutBody(url) 
  on error resume next
  Set Retrieval = CreateObject("Micro"&"soft.XML"&"HTTP") 
  With Retrieval 
  .Open "Get", url, False, "", "" 
  .Send 
  OutBody = .ResponseBody
  End With 
  Set Retrieval = Nothing 
End Function

Function OutToStr(body,Cset)
  dim objstream
  set objstream = Server.CreateObject("ado"&"db.stream")
  objstream.Type = 1
  objstream.Mode = 3
  objstream.Open
  objstream.Write body
  objstream.Position = 0
  objstream.Type = 2
  objstream.Charset = Cset
  OutToStr = objstream.ReadText 
  objstream.Close
  set objstream = nothing
End Function

Function Remote__Test()
s = "<td width='170' align='center'><a href='http://www.dgchr.com/'>"
s = s&"<img src='http://www.dgchr.com/img/img20/logo.jpg' width='135' height='86' /></a>"
s = s&"<a> http://www.dg.gd.cn/xad/2010/dgw.jpg </a>"
s = s&"<a> http://www.dgchr.com/script/sys/adpic/0BA85C57DC4FGV9U9GGXX1SW1.jpg </a>"
s = s&"<a> <img src=""http://www.dg.gd.cn/sys/pinc/pimg/dgnetlogo.gif"" /> </a>"
s = s&"<a> http://localhost:240/u/demo/pfile/pimg/qqimg21_13.gif </a>"
s = s&"<a> http://www.dgchr.com/img/img20/pic/pic_125.jpg </a>"
s = s&"<a> http://www.dgchr.com/img/img20/bar_05.jpg </a>" 

s = s&"<a> http://oa.dg.gd.cn/pimg/logo/logo.jpg </a>"
s = s&"<a>  </a>"

s = s&"</td>"
s = s&"<a> http://www.dgchr.com/img/img20/pic/pic_129.jpg </a>"
s = RemoteReplaceUrl(s, "/u/demo/upfile/", "defdt-2010") 'dtinf-2010-8Q-F42A.486FG
Response.Write Replace(s,"<a>","<br>")
End Function

'================================================
'作  用：替换字符串中的远程文件为本地文件并保存远程文件
'参  数：
'	sHTML		: 要替换的字符串
'	sExt		: 执行替换的扩展名
' s = RemoteReplaceUrl(sCont, "/u/demo/upfile/", "dtinf-2010-8Q-F42A.486FG")
' sAllowExt = "gif|jpg|jpeg"'oRs("S_ImageExt")
'================================================
Function RemoteReplaceUrl(sHTML, sPath, sKeyID)
	Dim sContentPath,s_Content,sImgs,sExt
	s_Content = sHTML
	sExt = "gif|jpg|jpeg"
	sContentPath = sPath&Replace(sKeyID,"-","/")&"/" ': Response.Write sContentPath
	If Obj_Test("Microsoft.XMLHTTP",false) = False then
		'Response.Write "<br>:er1"
		RemoteReplaceUrl = s_Content
		Exit Function
	End If
	'Response.Write "<br>:er2" 'OK
	Dim re, RemoteFile, RemoteFileurl, SaveFileName, SaveFileType
	Set re = new RegExp
	re.IgnoreCase  = True
	re.Global = True
	sImgs = Get_HLinks(s_Content,"<img[^<]*>") ' 蠢办法了... Peace加这行
	re.Pattern = "(http:(\/\/|\\\\){1}(([A-Za-z0-9_-])+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & sExt & ")))"
	'(http|https|ftp|rtsp|mms) : 保留http, 其它协议的不要...
	'Set RemoteFile = re.Execute(s_Content)  ' 蠢办法了... Peace不用这行，用下1行
	Set RemoteFile = re.Execute(sImgs)
	Dim a_RemoteUrl(), n, i, bRepeat
	n = 0
	' 转入无重复数据
	For Each RemoteFileurl in RemoteFile
		'Response.Write "<br>1:"&RemoteFileurl 'OK
		If n = 0 Then
			n = n + 1
			Redim a_RemoteUrl(n)
			a_RemoteUrl(n) = RemoteFileurl
		Else
			bRepeat = False
			For i = 1 To UBound(a_RemoteUrl)
				If UCase(RemoteFileurl) = UCase(a_RemoteUrl(i)) Then
					bRepeat = True
					Exit For
				End If
			Next
			If bRepeat = False Then
				n = n + 1
				Redim Preserve a_RemoteUrl(n)
				a_RemoteUrl(n) = RemoteFileurl
			End If
		End If		
	Next
	' 开始替换操作
	nFileNum = 0
	For i = 1 To n
		'Response.Write "<br>2:"&a_RemoteUrl(i) 'OK
		'Response.Write sPath&"-"&sKeyID&"-"&sContentPath
		Call fold_add9(sPath, sKeyID, 0)
		SaveFileType = Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), ".") + 1) 'GetRndFileName(SaveFileType)
		SaveFileName = Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",4)&"("&i&"."&SaveFileType
		If RemoteSaveFile(sContentPath&SaveFileName, a_RemoteUrl(i), "", "") = True Then
			
			nFileNum = nFileNum + 1
			If nFileNum > 0 Then
				sOriginalFileName = sOriginalFileName & "|"
				sSaveFileName = sSaveFileName & "|"
				sPathFileName = sPathFileName & "|"
			End If
			sOriginalFileName = sOriginalFileName & Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), "/") + 1)
			sSaveFileName = sSaveFileName & SaveFileName
			sPathFileName = sPathFileName & sContentPath & SaveFileName
			s_Content = Replace(s_Content, a_RemoteUrl(i), sContentPath&SaveFileName, 1, -1, 1)
		End If
	Next

	RemoteReplaceUrl = s_Content
End Function

'================================================
'作  用：保存远程的文件到本地
'参  数：s_LocalFileName ------ 本地文件名
'		 s_RemoteFileUrl ------ 远程文件URL
'返回值：True  ----成功
'        False ----失败
'<!-- From eWebEditor(RemoteSaveFile) --- Edit by Peace -->
'================================================
Function RemoteSaveFile(s_LocalFileName, s_RemoteFileUrl, n_AllowSize, s_AllowExt)
    Dim nAllowSize,sExt1,sExt2,sAllowExt,PageThis
	If n_AllowSize="" Then  '最大文件大小(KB) 
	  nAllowSize = 960
	Else
	  nAllowSize = n_AllowSize
	End If
	If s_AllowExt="" Then  '判断文件类型(.xml.htm.txt)
	  sAllowExt = "(.jpg.gif.jpeg.png.tif.swf.flv)"
	Else
	  sAllowExt = s_AllowExt
	End If
	sExt1 = Mid(s_LocalFileName,InStrRev(s_LocalFileName,"."),8)
	'sExt2 = Mid(s_RemoteFileUrl,InStrRev(s_RemoteFileUrl,"."),8) 'Or inStr(sAllowExt,sExt2)<=0 
	If inStr(sAllowExt,sExt1)<=0 Then
	  RemoteSaveFile = False
	  Exit Function
	End If '// 判断文件类型结束 =========================================== End
	PageThis = Request.ServerVariables("HTTP_HOST") ' 本页地址, 不含http:// 含端口; localhost:240
    'PageThis = "www.dg.gd.cn" '测试，或屏蔽某站图片
	If Left(s_RemoteFileUrl,Len(PageThis)+7)="http://"&PageThis Then ' 同一站点，不执行
	  RemoteSaveFile = False
	  Exit Function
	End If ' ============================================================== 同一站点，不执行 End
	Dim Ads, Retrieval, GetRemoteData
	Dim bError
	bError = False
	RemoteSaveFile = False
	'Response.Write s_LocalFileName
	'On Error Resume Next
	Set Retrieval = Server.CreateObject("Micro"&"soft.XML"&"HTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing

	If LenB(GetRemoteData) > nAllowSize*1024 Then
		bError = True
	Else
		Set Ads = Server.CreateObject("ado"&"db.stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile Server.MapPath(s_LocalFileName), 2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
	End If
    'Response.Write Err.Number
	If Err.Number = 0 And bError = False Then
		RemoteSaveFile = True
	Else
		Err.Clear
	End If
End Function



' 转为根路径格式
Function Path_Relative2Root(url)
	Dim sTempUrl
	sTempUrl = url
	If Left(sTempUrl, 1) = "/" Then
		Path_Relative2Root = sTempUrl
		Exit Function
	End If

	Dim sWebEditorPath
	sWebEditorPath = Request.ServerVariables("SCRIPT_NAME")
	sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Do While Left(sTempUrl, 3) = "../"
		sTempUrl = Mid(sTempUrl, 4)
		sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Loop
	Path_Relative2Root = sWebEditorPath & "/" & sTempUrl
End Function

' 根路径转为带域名全路径格式
Function Path_Root2Domain(url)
	Dim sHost, sPort
	sHost = Split(Request.ServerVariables("SERVER_PROTOCOL"), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
	sPort = Request.ServerVariables("SERVER_PORT")
	If sPort <> "80" Then
		sHost = sHost & ":" & sPort
	End If
	Path_Root2Domain = sHost & url
End Function


%>
