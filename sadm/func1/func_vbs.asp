<%

Function IIf(xExp, Val1, Val2)
  If xExp = True Then
    IIf = Val1
  Else
    IIf = Val2
  End If
End Function
Function echo(xStr)
  Dim s :s=xStr
    s = Replace(s,"<","&lt;")
    s = Replace(s,"[br]","<br/>")
	s = Replace(s,vbcrlf,"<br>"&vbcrlf)
    Response.Write vbcrlf&"<br>"&s
End Function

Function Get_BSize(xByte)
  Dim b
  b = Int(RequestSafe(xByte,"N","0"))
  If b>(1024^3) Then
    b = FormatNumber(b/(1024^3),2) &" (GB)"
   ElseIf b>(1024^2) Then
    b = FormatNumber(b/(1024^2),2) &" (MB)"
  ElseIf b>1024 Then
    b = FormatNumber(b/1024,2) &" (KB)"
  Else
    b = b &" (B)"
  End If
  Get_BSize = b
End Function

Function Get_CIP()
  Dim IP
  IP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  If IP = "" Then
    IP = Request.ServerVariables("REMOTE_ADDR")
  End If
  Get_CIP = IP
End Function

'// Response.Write Get_State("Y","N;Y","否;是")
Function Get_State(xState,xKey,xMsg)
Dim sc,ac,ak,am,j,r,i
  if xMsg="" Then xMsg=xKey
  sc="000;F00;0FF;00F;FF0;0F0;F0F;800;880;080;088;008;808;888" 
  ac=Split(sc,";") : ak=Split(xKey,";") : am=Split(xMsg,";")
  j=0 : r="<span style='color:#CCC'>-</span>"
  For i=0 To uBound(ak)
    If j>uBound(ac) Then j=0
    If xState=ak(i) Then 
      r = "<span style='color:#"&ac(j)&"'>"&am(i)&"</span>"
      Exit for
    End If
    j = j + 1
  Next
Get_State = r
End Function

Function Get_Option(xmid,xfirst,xend,xstep)
  Dim ofirst,oend,i,ostr
   ofirst = xfirst
   oend = xend
   If xmid = "" Then
     omid = xfirst
   End If  
   ostr = ""
   For i = xfirst To xend step xstep
      ostr = ostr & vbcr
      If i = xmid Then
       ostr = ostr & "<option value="""&i&""" selected>"&i&"</option>"
    Else    
       ostr = ostr & "<option value="""&i&""">"&i&"</option>"
    End if
   Next
  Get_Option = ostr 
End Function

Function Get_SOpt(xCode,xName,xVal,xFlag)
Dim ac,an,i,tv,tl,tx,fSel,str
 If xName="" Then xName=xCode
 ac = split(xCode,";")
 an = split(xName,";")
 str = ""
 txx = ""
 For i = 0 to uBound(ac)
   tv = ac(i)
   tl = an(i) 
   If xVal = tv Then
	   fSel = " selected "
	   If xFlag = "Val" Then 
	        txx = tl
			Exit For
	   End If
   Else
	   fSel = ""
   End If
   str = str&vbcrlf& "<option value="&tv&fSel&">"&tl&"</option>"
 Next
	   If xFlag = "Val" Then 
	        str = txx
	   Else
	        str = str       
	   End If
Get_SOpt = str
End Function


Function Check_RTest(xStr,xPatrn)
  Dim rExp, retVal ' 建立变量。 
  Set rExp = New RegExp ' 建立正则表达式。 
  'rExp.Global = True
  rExp.Pattern = xPatrn ' 设置模式。 
  rExp.IgnoreCase = true ' 设置是否区分大小写。 
  retVal = rExp.Test(xStr) ' 执行搜索测试。
  Set rExp = Nothing 
  If retVal Then 
    Check_RTest = true 
  Else 
    Check_RTest = false
  End If 
End Function
'Show_RExp(s3,"</P>|<.*?>","[BR]")
Function Show_RExp(xStr,xPatrn,xObj)
  Dim rExp,tStr : tStr=xStr            ' 建立变量。
  Set rExp = New RegExp                ' 建立正则表达式。
    rExp.Global = True                 ' 设置为全局
    rExp.Pattern = xPatrn              ' 设置模式。
    rExp.IgnoreCase = true             ' 设置是否区分大小写。
    Show_RExp = rExp.Replace(tStr,xObj) ' 作替换。
  Set rExp = Nothing
End Function

Function Show_sTitle(xText,xColor)
Dim oStr,tCol
  tCol = xColor :oStr = xText
  If inStr(tCol,"#")<=0 Then 
    tCol="#"&tCol
  End If
  oStr = Replace(oStr,chr(34),"'")
  oStr = Replace(oStr,"<","&lt;")
  oStr = Replace(oStr,">","&gt;")
  If Len(tCol)>3 And inStr(tCol,"000000")<=0 Then
    oStr = "<font color='"&tCol&"'>"&oStr&"</font>"
  End If
Show_sTitle = oStr
End Function

Function Show_SLen(xSubj,xLen)
Dim i,j,xStr,ch,s2
i=0 : j=0 : s2="" : xStr=xSubj&""
 For i=1 To Len(xStr)
  ch = Mid(xStr,i,1)
  j=j+0.5 '半个中文宽度,中文就再处理下一行
  If Asc(ch)<30 OR Asc(ch)>127 Then j=j+0.5 
  If j>xLen Then Exit For
  s2=s2&Mid(xStr,i,1)
 Next
 If xStr<>s2 Then 
   s2=Mid(s2,1,Len(s2)-2)&"..."
 End If
Show_SLen = Show_Text(s2)
End Function

Function Show_Html(xStr) ' <script type="text/javascript"> <iframe> <object> <applet> <!---->
  Dim xHtml : xHtml = xStr&""
   xHtml = Show_RExp(xHtml,"<script[^|]*</script>|<script[^<]*</script>|<script[^>]*</script>","")
   xHtml = Show_RExp(xHtml,"<iframe[^|]*</iframe>|<object[^|]*</object>|<applet[^|]*</applet>","")
   xHtml = Show_RExp(xHtml,"<script","&#60;script")
   xHtml = Show_RExp(xHtml,"<iframe","&#60;iframe")
   xHtml = Show_RExp(xHtml,"<object","&#60;object")
   xHtml = Show_RExp(xHtml,"<applet","&#60;applet")
   xHtml = Show_RExp(xHtml,"onmouse","&#79;nmouse")
   xHtml = Show_RExp(xHtml,"onload","&#79;nload")
   xHtml = Replace(xHtml, "<!--", "&#60;!--")
   xHtml = Replace(xHtml, "-->", "--&#62;")
   Show_Html = xHtml
End Function

Function Show_HView(xStr) ' Save,Edit,UBB,show text in html
  Dim xText : xText = xStr&""
   xText = Show_Text(xText)
   xText = Show_RExp(xText, "&lt;A ", "<A ")
   xText = Show_RExp(xText, "&lt;/A&gt;", "</A>")
   xText = Show_RExp(xText, "&lt;B&gt;", "<B>")
   xText = Show_RExp(xText, "&lt;/B&gt;", "</B>")
   xText = Show_RExp(xText, "&lt;I&gt;", "<I>")
   xText = Show_RExp(xText, "&lt;/I&gt;", "</I>")
   xText = Replace(xText, "/&gt;", "/>")
   xText = Show_RExp(xText,"OnMouse","&#79;nMouse") '/ OnMouseOver
   Show_HView = xText
End Function 

Function Show_HText(xStr,xLen) 
  Dim tStr : tStr=xStr
  tStr = Show_RExp(tStr,"<SCRIPT[^|]*</SCRIPT>|<SCRIPT[^<]*</SCRIPT>|<SCRIPT[^>]*</SCRIPT>","")
  tStr = Show_RExp(tStr,"<.*?>","")
  tStr = Replace(tStr,"&nbsp;"," ")
  tStr = Replace(tStr,vbcrlf,"") '以下4行去掉多余空白
  'tStr = Replace(tStr," ","")
  tStr = Replace(tStr,"  "," ")
  tStr = Replace(tStr,"　"," ")
  tStr = Left(tStr,xLen) 
  Show_HText = tStr
End Function

Function Show_Text(xStr) ' Save,Edit,UBB,show text in html
   Dim xText : xText=xStr
   xText = xText&""
   xText = Replace(xText, "<", "&lt;")
   xText = Replace(xText, ">", "&gt;")
   xText = Replace(xText, chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")    'TAB  
   xText = Replace(xText, "  ", "&nbsp;&nbsp;")     
   xText = Replace(xText, vbcrlf, "<br>")
   xText = Replace(xText, vbcr, "<br>")
   xText = Replace(xText, vblf, "<br>")
   Show_Text = xText
End Function 

Function  Show_Form(xStr) ' show in form items
   Dim xText : xText = xStr&""
   xText = Replace(xText, chr(34), "&#34;")
   xText = Replace(xText, "'", "&#39;") 
   xText = Replace(xText, "<", "&lt;")
   xText = Replace(xText, ">", "&gt;")
   Show_Form = xText
End Function 

Function  Show_jsStr(xStr) ' show in javascript string
   Dim xText : xText = xStr&""
   xText = Replace(xText, "\", "\\")    ' ()[]{},*+-,$?,\|,^'
   xText = Replace(xText, "/", "\/")    ' ()[]{},*+-,$?,\|,^'
   xText = Replace(xText, chr(34), "\"&chr(34))
   xText = Replace(xText, chr(39), "\"&chr(39))
   xText = Replace(xText, vbcrlf, "\r\n")
   xText = Replace(xText, vbcr, "\r")
   xText = Replace(xText, vblf, "\n")
   'xText = Replace(xText, "<", "&lt;")
   'xText = Replace(xText, ">", "&gt;")
   Show_jsStr = xText
End Function 

Function  Show_jsKey(xKey) ' show in javascript string
   Dim xText : xText = xKey&""
   xText = Replace(xText, " ", "")
   xText = Replace(xText, "-", "_")
   xText = Replace(xText, ".", "_")
   Show_jsKey = xText
End Function 

Function Get_vPath(xLen)
  Dim vHost,vPath
  if NOT IsNumeric(xLen) then
    xLen = 0
  end if
  vHost = Request.ServerVariables("SERVER_NAME")
  port = Request.ServerVariables("SERVER_PORT")
  If port<>"80" Then
    vHost = vHost&":"&port
  End If
  vPath = Request.ServerVariables("URL")
  vPath = Left(vPath,len(vPath)-xLen)
  Get_vPath = "http://"&vHost&vPath
End Function

Function Get_fName()
  Dim sName
  sName = Request.ServerVariables("Script_Name") 'Request.ServerVariables("URL")
  Get_fName	 = StrReverse(Mid(StrReverse(sName),1,InStr(1,StrReverse(sName),"/")-1))
End Function


Function RequestF(xPName,xPType,xLen) 
  Dim tVal : tVal=Request.Form(xPName)
  RequestF = RequestSafe(tVal,xPType,xLen)
End Function 
Function RequestQ(xPName,xPType,xDefault) 
  Dim tVal : tVal=Request.QueryString(xPName)
  RequestQ = RequestSafe(tVal,xPType,xDefault)
End Function 
Function RequestS(xPName,xPType,xDefault) 
  Dim tVal : tVal=Request(xPName)
  RequestS = RequestSafe(tVal,xPType,xDefault)
End Function 
Function RequestSafe(xPName,xPType,xDefault) 
  Dim PValue 
  PValue = xPName&""
  If xPType = "N" then ' N/C/D/ID/PW/EM
	If not isNumeric(PValue) then 
      PValue = xDefault ' -1;0;1
    End if 
  ElseIf xPType = "D" then 
	if ( NOT isDate(PValue) ) AND ( isDate(xDefault) ) then
      PValue = xDefault
    elseif ( NOT isDate(PValue) ) AND ( NOT isDate(xDefault) ) then 
      PValue = "1900-12-31"
	end if
  ELSE  'C Save
	PValue = LeftB(PValue,xDefault)
	PValue = Replace(PValue, "'", "''")		
  End if 
RequestSafe = PValue
End Function 


Function js_Alert(xMsg,xAct,xAddr) 
  Dim StrJS
  StrJS = vbcr& "<script type='text/javascript'>"
  StrJS = StrJS& vbcr& "alert('"&xMsg&"');"
  if xAct = "Back" then
    StrJS = StrJS& vbcr& "history.go("&xAddr&");"
  elseif xAct = "Close" then
    StrJS = StrJS& vbcr& "window.close();"
  elseif xAct = "Open" then
    StrJS = StrJS& vbcr& "window.open('"&xAddr&"');"
  elseif xAct = "Redir" then
    StrJS = StrJS& vbcr& "location.href='"&xAddr&"';"
  elseif xAct = "" then
    '//
  else 'Alert
    StrJS = StrJS& vbcr& "history.go("&xAddr&");"
  end if
  StrJS = StrJS& vbcr& "</script>"
  js_Alert = StrJS
End Function


Function Chr_Filter(xObjStr)
  Dim sKey,oStr,aKey,i,iKey,fKey : fKey = ""
  sKey = ParFilALLKeys
  oStr = xObjStr
  sKey = Replace(sKey,",",";")
  aKey = Split(sKey,";")
  For i = 0 To uBound(aKey)
    iKey = Trim(aKey(i)&"")
    If (Len(iKey)>2) And (inStr(oStr,iKey)>0) Then
	    fKey = iKey&"," 
		Exit For
    End If
  Next
  Chr_Filter = fKey
End Function

Function Chr_Fil2(xObjStr)
 Dim sKey,oStr,aKey,i,iKey
 sKey = ParFilALLKeys
 oStr = xObjStr
  sKey = Replace(sKey,",",";")
  aKey = Split(sKey,";")
  For i = 0 To uBound(aKey)
    iKey = Trim(aKey(i)&"")
    If (LenB(iKey&"")>2) And (inStr(oStr,iKey)>0) Then
	  oStr = Replace(oStr,iKey,String(Len(iKey&""),"*")) 
    End If
  Next
  Chr_Fil2 = oStr 
End Function

%>

