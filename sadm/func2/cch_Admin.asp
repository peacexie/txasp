<!--#include file="func_const.asp"-->
<!--#include file="cch_Class.asp"-->
| <a href='?'>Null</a> 
| <a href='?~act=~Test'>~Test</a> 
| <a href='?act=Clear'>Clear</a> 
| <a href='?act=Update'>Update()</a> 
| <a href='?act=View'>View</a> 
| <a href="cch_Demo.asp">Demo</a> | <br />
| <a href='?act=<%=Config_Path%>ext/link.asp'>/ext/link.asp - RAM</a> 
| <a href='?f=<%=Config_Path%>ext/map.asp'>/ext/map.asp - File</a> 
| 
<hr>
<%

'使用：配置好参数 Config_Path
'Call cchCheck("") ''配置以上参数
'Call cchFile(Config_Path&"ext/map.asp")
'Call cchRAM(Config_Path&"ext/link.asp")

' Call fileSave(xPName2,xCont,xCSet)
' f = fileExist(pathName)
' s = fileRead(pathName,xCSet)
' url = urlPath(sPath)


act = Request("act")
If act = "Clear" Then
  Application.Contents.Removeall()
  Response.Write "Clear: Clear Cache OK!"
ElseIf act = "Update" Then
  Call cchClear()
  Response.Write "cchClear(): Clear Cache OK!"
ElseIf Request("~act")="~Test" Then
  Response.Write Request("~act")
  Response.Write "<br>URL:"&Request.ServerVariables("URL")
  Response.Write "<br>Query:"&Request.QueryString()
  Response.Write "<br>"&Request.ServerVariables("SERVER_NAME")
  'Response.End()
  Response.Write urlPath(Config_Path&"ext/duty.asp")&"<br>"
  'Call urlSave(Config_Path&"ext/map.asp",Config_Path&"upfile/sys/cache/Cache[~ext~map.asp].urlSave.htm") 
  Response.Write  Int("987654321012")&"<br>"
  Response.Write CInt("12345")&" CInt<br>"
  Response.Write CInt("32767")&" CInt<br>"
  Response.Write CLng("1234567890")&" CLng<br>"
  Response.Write CLng("2147483647")&" CLng<br>"
  Response.Write CLng("2140123456")&" CLng<br>"
  Response.Write Date()&" Date<br>"
  Response.Write Time()&" Time<br>"
  Response.Write Now()&" Now<br>"
  Response.Write strRand()&"<br>"
  'Response.Write cchEnd("strCont","Upd/Read","cchRAM/cchFile")&"<br>" ''End()
  ' Call cchClear()
ElseIf act = "View" Then
  For Each Key in Application.Contents
	sCont = application(Key)
	If Len(sCont)>128 Then
	  sCont = Server.HTMLEncode(Left(sCont,96)&" ……………… "&Right(sCont,64))
	End If
	Response.Write vbcrlf&vbcrlf&"<br><br><hr>"&vbcrlf&"("&Key&"):"&sCont
  Next
ElseIf Request("f") <> "" Then
  Call cchFile(Request("f"))
ElseIf act <> "" Then
  Call cchRAM(act)
Else
  Response.Write "Null Action: act=Clear,View Or [Path(Config_Path&/page/index.asp)]"
End IF

%>
