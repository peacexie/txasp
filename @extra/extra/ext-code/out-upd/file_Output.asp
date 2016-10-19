

<%

Function rs_GetFile(xKeyID,xImgPath,xSet) 
Dim KeyID

InfShow = ""
rs.Open "SELECT * FROM [NewsInfo] WHERE KeyRe='"&xKeyID&"' ORDER BY InfPNow ",conn,1,1 
  if NOT rs.eof then 
  i7 = 0
  Do While Not rs.EOF
  i7 = i7 + 1
KeyID = rs("KeyID")
KeyRe = rs("KeyRe")
InfPage = rs("InfPage")
InfPNow = rs("InfPNow")
If i7 = 1 Then
KeyMod = rs("KeyMod")
InfType = rs("InfType")
'TypName = rs_Val(conn,"SELECT TypName FROM NewsType WHERE TypLayer='"&InfType&"'","TypName")
KeyFlag = rs("KeyFlag")
InfSubject = Show_Text(rs("InfSubject"))
InfKey = Show_Text(rs("InfKey"))
InfFrom = Show_Text(rs("InfFrom"))
SetRead = rs("SetRead")
LogATime = rs("LogATime")
LogETime = rs("LogETime")
End If
InfContent = rs("InfContent")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
SetAdv = rs("SetAdv")
ImgName = rs("ImgName")&""
ImgAlign = rs("ImgAlign")
ImgTitle = rs("ImgTitle")
ImgWidth = rs("ImgWidth")
ImgHeight = rs("ImgHeight")
ImgScale = rs("ImgScale")
LogAddIP = rs("LogAddIP")
  imgStr = ""
If ImgName<>"" then
  imgStr = show_Media(xImgPath&ImgName,ImgWidth,ImgHeight,500,400)
ElseIf ImgName="" AND ImgTitle<>"" THEN
  imgStr = show_MPost(ImgTitle,ImgWidth,ImgHeight)
End If
InfContent = show_MPlace(InfContent,ImgAlign,imgStr,5)
InfShow = InfShow &vbcrlf&vbcrlf&"<p>"& InfContent&"</p>"
  rs.MoveNext
  Loop
  end if 
rs.Close()

 filName = InfKey&".htm"
 filData = InfShow
 filName = filName
 
 'Response.Write filName
 Call fil_add("/xxxxx/news/about/",filName,InfShow)  

End Function

%>


