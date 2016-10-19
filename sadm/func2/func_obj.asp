<%


'================================================
' *** 组件相关函数
'================================================
'检查组件是否被支持及组件版本的子程序
Function Obj_Version(sObj)
  Dim sVer : sVer=""
  If Obj_Test(strClassString,strEndFlag) Then
    sVer = TestObj.version
    if sVer="" or isnull(sVer) then sVer =TestObj.about
  Else
    sVer = "( ------ unKnow ------ )"
  End If
  Obj_Version = sVer 
End Function

'================================================
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'参  数：strEndFlag     ----如果没有安装，是否Response.End()
'返回值：True  ----已经安装
'        False ----没有安装
'================================================
Function Obj_Test(strClassString,strEndFlag)
	On Error Resume Next
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then 'If -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
	  Obj_Test = True'False'
	Else
	  Obj_Test = False
	  'If strEndFlag="" Then strEndFlag=false
	  If strEndFlag Then '////////// Show Msg and End
        set xTestObj = nothing
	    Response.Write "Error! "&strClassString&" NOT install in this Server! "
	    Response.End()
	  End If '//////////////////////
	End If
	Set xTestObj = Nothing
	Err = 0
End Function



'================================================
' *** 执行 cmd 命令！
'================================================
'启用：W.Script.S.hell，这里不检查Object了
'Do_Cmd("ping www.dg.gd.cn -w 1 -n 1")
Function Do_Cmd(xStr)
  Response.Buffer = true
  Set sObj = CreateObject("W"&"Script.S"&"hell") 
  Set cObj = sObj.Exec(xStr) 
  sRes = cObj.StdOut.Readall() 
  Set cObj = Nothing 
  Set sObj = Nothing 
  Do_Cmd = sRes
End Function



'================================================
' *** 水印相关
'================================================
' 配置区: /////////////////////////////////////////
'ConfigWMark = "T"                    ' T/P/N 文字/图片/无水印  动态设置
'Dim wMarkText,wMarkLogo        
'wMarkText = "[XXX测试网站]|domain.com" ' 水印文字:[XXX测试网站]|domain.com；|后面可以不要
'wMarkLogo = "upfile/myfile/logo/logo-peace.gif" ' 已知：Config_Path路径参数
' 调用：Call ImgPMark(pImg) ImgPMark(xOImg,"P")
' 调用：Call ImgTMark(pImg) ImgTMark(xOImg,"T")
' 注意要装好 Persits.Jpeg组件 //////////////////////
' 测试 Persits.Jpeg组件 Response.Write Obj_Test("Persits.Jpeg",true)
Function ImgMMark(xOImg,xFlag)

  If NOT fil_exist(xOImg) Then 
    Exit Function
  End If
  fExt = uCase(Mid(xOImg,InStrRev(xOImg,"."),12))
  If inStr("(.JPG.JPEG.GIF)",fExt)<=0 Then 
    Exit Function
  End If

  If xFlag="T" Then
    Call Obj_Test("Persits.Jpeg",true)
	Call ImgTMark(xOImg)
  ElseIf xFlag="P" Then 
    Call Obj_Test("Persits.Jpeg",true)
	Call ImgPMark(xOImg)
  Else
    
  End If
End Function 

Function ImgPMark(xOImg) ' 图片水印
  Dim fExt,mPhoto,mLogo,jPhoto,jLogo,xPos,yPos
  
  mPhoto=xOImg : mPhoto = Replace(mPhoto,"//","/")
  If instr(mPhoto,":")<=0 Then mPhoto=Server.MapPath(mPhoto) 
  mLogo=Config_Path&wmk_Logo 'xMImg  : mLogo = Replace(mLogo,"//","/")
  If instr(mLogo,":")<=0 Then mLogo=Server.MapPath(mLogo) 
  
  Set jPhoto = Server.CreateObject("Persits.Jpeg")
  jPhoto.Open mPhoto
  Set jLogo = Server.CreateObject("Persits.Jpeg")
  jLogo.Open mLogo

  Dim ma :ma = ImgWMPos(jPhoto.OriginalWidth,jPhoto.OriginalHeight,jLogo.OriginalWidth,jLogo.OriginalHeight)
  If ma(0) Then
    xPos = ma(1) :yPos = ma(2)
	'jPhoto.DrawImage xPos,yPos,jLogo,1,&HFFFFFF,wmk_TScope 
	jPhoto.DrawImage xPos,yPos,jLogo,wmk_Trans,wmk_TColor
	'第一,二个参数是X,Y offsets,第三个参数是水印图片;
	'第四个参数是水印透明度,第5个参数是抽取透明层的颜色; (第4个到第六个参数是可选参数)
	'第六个参数是抽取颜色的误差范围,有时候透明有杂点,可以通过这个参数调整.
	jPhoto.Save mPhoto
  End If
  
  Set jPhoto = Nothing
  Set jLogo = Nothing
End Function

Function ImgTMark(xOImg) ' 文字水印 
  Dim mPhoto,Jpeg,xPos,yPos,xText,aText,bText,cText,yOSet

  xText = wmk_Text
  If inStr(xText,"|")>0 Then
    aText = Split(xText,"|")
	bText = aText(0)
	cText = aText(1) 
	yOSet = wmk_HLine
  Else
	bText = xText
	cText = "" 
	yOSet = 0
  End If
  If inStr(wmk_Font,"|")>0 Then
	aFont = Split(wmk_Font,"|")
	wmk_Font1 = aFont(0)
	wmk_Font2 = aFont(1)
  Else
	wmk_Font1 = wmk_Font
	wmk_Font2 = wmk_Font
  End If

  mPhoto=xOImg : mPhoto = Replace(mPhoto,"//","/")
  If instr(mPhoto,":")<=0 Then mPhoto=Server.MapPath(mPhoto) 

  Set Jpeg = Server.CreateObject("Persits.Jpeg")
  Jpeg.Open mPhoto

  Dim ma :ma = ImgWMPos(Jpeg.OriginalWidth,Jpeg.OriginalHeight,wmk_Width,wmk_Height)
  If ma(0) Then
	Dim iArr,i,j,k
	iArr = Split("1,-1,-1,1,0",",")  
	jArr = Split("1,1,-1,-1,0",",")
	For k=1 To 5 '先4个象限移动1px,再到0,实现描边(Peace)
	  i = Int(iArr(k-1))
	  j = Int(jArr(k-1))
	  If k=5 Then 
	    iCol = wmk_Color1 '51文字颜色
	  Else
	    iCol = wmk_Color2 '51描边颜色
	  End If
	  Jpeg.Canvas.Font.Color = iCol             ' red 颜色&H0000FF
	  Jpeg.Canvas.Font.Size = wmk_Size               ' 大小 24
	  Jpeg.Canvas.Font.Bold = True                   ' 是否加粗
	  xPos = ma(1) :yPos = ma(2)
	  If cText="" Then
	    Jpeg.Canvas.Font.Family = wmk_Font1                ' 字体 Arial,Courier New
		Jpeg.Canvas.Print xPos+i, yPos+j, bText            ' 打印坐标x,y,需要打印的字符
	  Else
		Jpeg.Canvas.Font.Family = wmk_Font1
		Jpeg.Canvas.Print xPos+i, yPos+j, bText            ' 打印坐标x,y,需要打印的字符
		Jpeg.Canvas.Font.Family = wmk_Font2
		Jpeg.Canvas.Print xPos+i, yPos+j+yOSet, cText      ' 打印坐标x,y,需要打印的字符
	  End If
	Next

  End If
  
  '以下是对图片进行边框处理
  'Jpeg.Canvas.Pen.Color = &HFFFFFF              ' black 颜色
  'Jpeg.Canvas.Pen.Width = 2                     ' 画笔宽度
  'Jpeg.Canvas.Brush.Solid = False               ' 是否加粗处理
  'Jpeg.Canvas.Bar 1, 1, Jpeg.Width, Jpeg.Height ' 起始X坐标 起始Y坐标 输入长度 输入高度
  
  Jpeg.Save mPhoto
  Set Jpeg = Nothing
End Function

Function ImgWMPos(xW1,xH1,xW2,xH2) '获得水印位置
  Dim f,x,y,a(2)
  If xW1<xW2+60 Or xH1<xH2+60 Then
    a(0) = false
	a(1) = 1
	a(2) = 1
  Else
    a(0) = true
	If wmk_Pos="1" Then 
	  x = wmk_Padding
	  y = wmk_Padding
	ElseIf wmk_Pos="2" Then 
	  x = xW1-xW2-wmk_Padding
	  y = wmk_Padding
	ElseIf wmk_Pos="3" Then 
	  x = wmk_Padding
	  y = xH1-xH2-wmk_Padding
	ElseIf wmk_Pos="4" Then 
	  x = xW1-xW2-wmk_Padding
	  y = xH1-xH2-wmk_Padding
	Else
	  x = Int((xW1-xW2)/2)
	  y = Int((xH1-xH2)/2)
	End If
	a(1) = x
	a(2) = y
  End If
  ImgWMPos = a
End Function

'================================================
' *** 发邮件相关 '可选JmMethod=JMail,CDONTS,CDO
' 测试 JMAIL.Message组件 Response.Write Obj_Test("JMAIL.Message",true)
'================================================
Function Send_Mails(xFrom,xTo,xSubj,xBody,xCSet)   
  If JmMethod="CDO" Then
    f = Obj_Test("CDO.Message",false) 
	If f Then
	  Call Send_cdoMail(xFrom,xTo,xSubj,xBody,xCSet)
	End If
  ElseIf JmMethod="CDONTS" Then
	f = Obj_Test("CDONTS.NewMail",false) 
	If f Then
      Call Send_ntsMail(xFrom,xTo,xSubj,xBody,xCSet)
	End If
  Else 'jMail
    f = Obj_Test("JMAIL.Message",false)
	If f Then
	  Call Send_jMail(xFrom,xTo,xSubj,xBody,xCSet)
	End If
  End If
End Function

Function Send_ntsMail(xFrom,xTo,xSubj,xBody,xCSet)
  Set objMail = Server.CreateObject("CDONTS.NewMail")
  objMail.To = xTo ' "iampeace@21cn.com;iampeace2@263.net"
  'objMail.CC = JmAdmCC
  'objMail.Bcc = "iampeace@263.net;peace_2@sina.com" '暗送
  objMail.From = xFrom '"iampeace@263.net"
  objMail.Subject = xSubj
  objMail.Body = xBody
  'objMail.BodyFormat = 0 'Html
  'objMail.MailFormat = 0 '说明是以MIME发送
  'objMail.AttachFile("\\server\schedule\sched.xls", "SCHED.XLS") 
  'objMail.Importance = 2 '重要性 0-1-2:低-一般(默认)-高
  objMail.Send
  Set objMail = Nothing
End Function
Function Send_cdoMail(xFrom,xTo,xSubj,xBody,xCSet)  
  Dim cdoBaseUrl,objConfig,objMessage,objFields 
  cdoBaseUrl = "http://schemas.microsoft.com/cdo/configuration/"
  Set objConfig = Server.CreateObject("CDO.Configuration") 
  Set objFields = objConfig.Fields 
  With objFields 
	.Item(cdoBaseUrl&"sendusing")             = 2
	.Item(cdoBaseUrl&"smtpserver")            = JmailServer '"smtp.gmail.com" '邮件服务器
	.Item(cdoBaseUrl&"smtpserverport")        = 25 '465 '25 
	.Item(cdoBaseUrl&"smtpconnectiontimeout") = 10 
	.Item(cdoBaseUrl&"smtpauthenticate")      = 1 '0不需要验证, 1Basic验证方式, 2NTLM验证方式
	'.Item(cdoBaseUrl&"smtpusessl")            = true '可选
	.Item(cdoBaseUrl&"sendusername")          = JmailID '"xxx@gmail.com" '用户名 
	.Item(cdoBaseUrl&"sendpassword")          = JmailPW '"xxx" '密码 
	.Update 
  End With 
  Set objFields = Nothing 
  Set objMessage = Server.CreateObject("CDO.Message") 
  Set objMessage.Configuration = objConfig 
  Set objConfig = Nothing 
  With objMessage 
	.To = xTo '"elifebike@gmail.com"
	.CC = JmAdmCC '"xpigeon@163.com" 'Trim(Request.Form("m_user")) '接收者
	.From = """"&xFrom&"""<"&JmailID&">" '发送人（要和上面的邮件系统相同） 
	.Subject = xSubj '"Message-15 "&Now() ' 标题 
	'.TextBody = "Message-15 "&Now() ' 正文 
	.HTMLBody = xBody '"<h1>This is a message.</h1>" 
	'.CreateMHTMLBody "http://www.w3schools.com/asp/" 
	'.CreateMHTMLBody "file://c:/mydocuments/test.htm" 
	'.AddAttachment "F:\Work\peace_asp\code.96327.cn\img\Output.gif"'邮件附件
	'.BodyPart.Charset= "gb2312 "; 
	'.HTMLBodyPart.Charset= "gb2312 ";
	.Send 
  End With 
  Set objMessage = Nothing
End Function
Function Send_jMail(xFrom,xTo,xSubj,xBody,xCSet)         
  Dim jmail,jsilent
  jsilent=true 'false
  Set jmail = Server.CreateObject("JMAIL.Message")               'Jmail4.3
  jmail.silent          = jsilent        
  jmail.logging         = true          
  jmail.Charset                 = xCSet    '' 邮件的字符集，缺省为"US-ASCII","utf-8","GB2312"     
  jmail.ContentTransferEncoding = "base64" '' 内容传送时的编码方式，缺省是"Quoted-Printable"
  jmail.Encoding                = "base64" '' 附件编码方式
  jmail.ISOEncodeHeaders        = false    '' 头信息代码按照iso-8859-1字符设置，默认为True
  'xSubj = Send_EscCode(xSubj)
  xBody = Send_EscCode(xBody)
  jmail.ContentType             = "text/html"     
  jmail.AddRecipient    xTo 'JmAdmTo                           
  jmail.AddRecipientCC  JmAdmCC              
  jmail.Priority        = 3             
  jmail.Subject         = xSubj        
  jmail.Body            = xBody     
  'jmail.AddAttachment(Server.MapPath("01.jpg"),false,"image/jpg");
  JMail.FromName             = xFrom        'xFrom 
  jmail.From                 = JmailAccount 'emJmailID
  jmail.MailServerUserName   = JmailID      'emJmailID
  jmail.MailServerPassword   = JmailPW      'emJmailPW    
  jmail.Send(JmailServer)                   'emJSMTP  
  'Response.End()                
  jmail.Close   
  Set jmail = Nothing                         
End Function

Function Send_HCont(xField,xStart,xEnd)
Dim hCont,p1,p2,sKill
  hCont = RequestS(xField,"C",240000) 
  p1 = inStr(hCont,xStart)+Len(xStart)
  p2 = inStr(hCont,xEnd)
  If p1>24 and p2>p1 Then
    sKill = Mid(hCont,p1,p2-p1)
	hCont = Replace(hCont,sKill,"") 
  End If
  Send_HCont = hCont
End Function

Function Send_EscCode(p_Message)
  Dim m_char,m_asc,m_hex '字符，ASC码，16进制ASCII码
  Dim m_temp '临时字符
  Dim a_arc() 'ASC码数组
  Dim i 
  ReDim a_arc(Len(p_Message))
  For i = 0 To Len(p_Message) -1
    m_char = Mid(p_Message,i+1,1)
    m_asc = AscW(m_char) 
	If m_asc < 0 Then
	  m_temp = m_asc 
	  m_temp = m_temp+65536
      a_arc(i) = "&#"&m_temp&";"
	ElseIf m_asc < 32 Then
	  a_arc(i) = ""
	ElseIf m_asc < 255 Then
      a_arc(i) = m_char
	Else
	  m_temp = m_asc 
      a_arc(i) = "&#"&m_temp&";"
	End If
  Next
  Send_EscCode = Join(a_arc,"")
End Function
'Response.Write Send_EscCode("abcd测试谢永顺")

Function Send_T01()         
Dim jmail
Set jmail = Server.CreateObject("JMAIL.Message")               'Jmail4.3
  jmail.silent          = false 'true           
  jmail.logging         = true          
  jmail.Charset                 = "GB2312"    '' 邮件的字符集，缺省为"US-ASCII","utf-8","GB2312"     
  jmail.ContentTransferEncoding = "base64" '' 内容传送时的编码方式，缺省是"Quoted-Printable"
  jmail.Encoding                = "base64" '' 附件编码方式
  jmail.ISOEncodeHeaders        = false    '' 头信息代码按照iso-8859-1字符设置，默认为True
  xBody = Send_EscCode(xBody)
  jmail.ContentType             = "text/html"     
  jmail.AddRecipient    "dgweb@dg.gd.cn" 'JmAdmTo                                        
  jmail.Priority        = 3             
  jmail.Subject         = "通知13"&Now()        
  jmail.Body            = "开会13"&Now()  
  JMail.FromName             = "FromName#CompanyName"        'xFrom 
  jmail.From                 = "dgweb@dg.gd.cn"  'emJmailID From#unKnow
  jmail.MailServerUserName   = "dgweb@dg.gd.cn"  'emJmailID
  jmail.MailServerPassword   = "dgweb"      'emJmailPW    
  jmail.Send("mail.dg.gd.cn")                   'emJSMTP   
  Response.Write "T01:send OK!"            
jmail.Close   
Set jmail = Nothing   

    Response.Write jsilent&":<br>"&xCSet&":<hr>"
	Response.Write xTo&":<br>"&JmAdmCC&":<br>"&xFrom&":<hr>"
    Response.Write JmailAccount&":<br>"&JmailID&":<br>"&JmailPW&":<br>"&JmailServer&":<hr>"
    Response.Write xSubj&":<br>"&xBody&":<br>"
                      
End Function
'Call Send_T01() 



'================================================
' *** 自动缩略图 相关
'================================================
' 测试 Response.Write ImgSUpd("InfoNews","xKey",InfCont,120,60)
Function ImgSmall(xUrl,xW,xH) ' 生成缩略图
  Dim nW,nH '为空 / '文件不存在
  If NOT Obj_Test("Persits.Jpeg",False) Then '组件未安装
    ImgSmall = "" 'xUrl 
	Exit Function
  End If
  Set Jpeg = Server.CreateObject("Persits.Jpeg") '调用组件
  Jpeg.Open Server.MapPath(xUrl)'打开图片
  nW=Jpeg.OriginalWidth : nH=Jpeg.OriginalHeight
  If nW>xW Then
	nH=Int(xW*nH/nW) :nW=xW
  End If
  If nH>xH Then
	nW=Int(xH*nW/nH) :nH=xH
  End If
  Jpeg.Width = nW 'Jpeg.OriginalWidth / 2 '高与宽为原图片的1/2
  Jpeg.Height = nH 'Jpeg.OriginalHeight / 2 '保存图片
  Jpeg.Save Server.MapPath(xUrl)
  Set Jpeg = Nothing
  ImgSmall = xUrl
End Function

Function ImgSUrl(xData) ' 得到缩略图Url
  Dim s,a,url,f,t,e,i
  s = Get_HLinks(xData,"<img[^<]*>")
  a = Split(s,"||") : url = "" : f = ""
  For i=0 To uBound(a)
  If a(i)<>"" Then
	If inStr(a(i),"/upfile/")>0 Then '在附件目录
	  t = Get_1Url(a(i),"src=")
	  If t<>"" And Left(t,1)="/" And fil_exist(t) Then '以/开始,文件存在
	  e = lCase(Mid(t,InStrRev(t,"."),8))
	  If inStr("(.jpeg.jpg.gif)",e)>0 Then '是一个图片后缀
	    url = t
	    Exit For
	  End If
	  End If
	End If
  End If
  Next
  ImgSUrl = url
End Function

Function ImgSUpd(xTab,xKey,xCont,xW,xH) ' 保存缩略图
  Dim url,n,p,f,flag,sql : f = "" ': flag = true
  url = ImgSUrl(xCont) ': Response.Write url
  If url<>"" Then
	n = InStrRev(url,"/") : p = Left(url,n)
	f = Mid(url,n+1)      : f = Replace(f,".","~copy.")
	Call fil_copy(url,p&f)
	xW = Int(RequestSafe(xW,"N",0))
	xH = Int(RequestSafe(xH,"N",0))
	If xW>0 And xH>0 Then
	  Call ImgSmall(p&f,xW,xH) 
	ElseIf xW>0 Then
	  Call ImgSmall(p&f,xW,9999) 
	ElseIf xH>0 Then
	  Call ImgSmall(p&f,9999,xH) 
	End If 
	sql = "UPDATE ["&xTab&"] SET ImgName='"&"^"&f&"^^^^^^^^"&"' WHERE KeyID='"&xKey&"'"
	Call rs_Dosql(conn,sql)
  End If
  ImgSUpd = f 'ImgName = this()
End Function
'ModX888ImgSCopy = "Y"
'ModX888ImgSAtuo = "120x90"
'Dim PicR124ImgSCopy,PicR124ImgSAtuo : PicR124ImgSCopy = "Y" : PicR124ImgSAtuo = "90x120"

%>

