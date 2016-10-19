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
  InfType = rs("InfType")
  InfName = rs("InfName")
  InfCont = rs("InfCont")
  InfPara = rs("InfPara")
  end if 
  rs.Close()
  SET rs=Nothing 
  
  sCode = gCodeStr(InfCont,InfType,InfPara)
  
  sCode = Replace(sCode,"(sPath)",Config_Path)
  Response.Write "<font style='font-size:18px;'>调用代码:<br> "
  Response.Write Server.HTMLEncode("<script src='"&Config_Path&"upfile/sys/xadv/txt_"&InfType&".js'></script>")
  Response.Write "<br>效果:<hr>"&sCode&""
  Response.Write "</font>"

Else

  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open "SELECT * FROM [WebAdvert] WHERE KeyMod='AdText'",conn,1,1 
  if NOT rs.eof then 
  Do While Not rs.EOF

  InfType = rs("InfType")
  InfName = rs("InfName")
  InfCont = rs("InfCont")
  InfPara = rs("InfPara")
  
  sCode = gCodeStr(InfCont,InfType,InfPara)
  'Response.Write sCode
  
  rs.MoveNext()
  Loop
  end if 
  rs.Close()
  SET rs=Nothing 
  
End If


Function gCodeStr(xCont,xType,xPara) 
  Dim s,i,j
  ParaA = Split(InfPara,"|")
  Para1 = ParaA(0) : Para2 = ParaA(1) 
  s = "<div id='' style='width:"&Para1&"px; height:"&Para2&"px;'>"&xCont&"</div>"
  'Call File_Add2("../../upfile/myfile/xadv/txt_"&xType&".htm",sData,"UTF-8")
  sData = Replace(s,"(Config_Path)",Config_Path)
  Call File_AddJS(xType,sData)
  gCodeStr = s
End Function

Function File_AddJS(xType,xData)
  Dim s
  s = xData
  s = Replace(s,"\","\\")
  s = Replace(s,"'","\'")
  s = Replace(s,"""","\""")
  a = Split(s,vbcrlf)
  s = ""
  for i=0 To uBound(a)
  s = s&vbcrlf&"document.write("""&a(i)&""");"
  Next
  Call File_Add2(Config_Path&"upfile/sys/xadv/txt_"&xType&".js",s,"UTF-8")
End Function

%>

</body>
</html>
