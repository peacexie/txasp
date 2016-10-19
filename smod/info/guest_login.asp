<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../inc/form/form_app.asp"-->

<%
act = Request("act") 'in,out
If act = "time" Then

 Response.Write vbcrlf&"getElmID('"&Request("TimID")&"').value = '"&Get_AMin(0)&"';" 
 Response.Write vbcrlf&"getElmID('fmGuestLogin').submit();"
 Response.End()
 
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Guest Login...</title>
</head>
<body>
<%If SwhGuestPub<>"Y" Then%>
<table width="480" border="0" align="center">
  <tr>
    <td>提示：禁止游客发布</td>
  </tr>
</table>
<%Else%>
<%
If act = "in" Then
  
  ModID = RequestF("ModID","C",48)    'InfN128
  TypID = RequestF("TypID","C",48)    'N111012;
  ChkSes = RequestF("ChkSes","C",48)
  ChkUrl = "http://"&RequestF("ChkUrl","C",255)    'http://localhost:315/page/info.asp
  ChkA29 = RequestF(App29Code,"C",255) '2EDFD61
  
  cSes = Hex(Session.SessionID)
  cUrl = Request.Servervariables("HTTP_REFERER") 'http://localhost:315/page/info.asp?ModID=InfN128&TypID=N111012
  cApp29 = Get_AChk(ChkA29,1)
  
  If cSes<>ChkSes Then
    Response.Write "错误：(cSes)可能超时,刷新后再提交!" '我要上传
	Response.End()
  ElseIf inStr(cUrl,ChkUrl)<=0 Then
    Response.Write "错误：(cUrl)不允许外部提交!"
	Response.End()
  ElseIf Not cApp29 Then
    Response.Write "错误：(cTime)可能超时,刷新后再提交!"
	Response.End()
  Else

	Call Add_Log(conn,Session.SessionID,"guest登入","[guest]","")
	
	Response.Write "<br>"&ModID
	Response.Write "<br>"&TypID
	Response.Write "<br>"&ChkSes
	Response.Write "<br>"&ChkUrl
	Response.Write "<br>"&cUrl
	Response.Write "<br>"&ChkA29
	
	Session("MemID") = "guest"
	Session("MemPerm") = "{ (place) }"&"{gMemPass};("&ModID&");" '(MemFEditor);(MemFUpload)
	url = "info_add.asp?ModID="&ModID&"&InfType="&TypID&"&PrmFlag=(Mem)"
	Response.Redirect url
	
	Response.Write "<br>"&Session("MemID")
	Response.Write "<br>"&Session("MemPerm")
	Response.Write "<br><a href='"&url&"'>Publish</a>"

  End If 

ElseIf act = "out" Then 

  Session("MemID") = ""
  Session("MemPerm") = ""
  Response.Write js_Alert(Request("msg")&"\n 关闭窗口!","Close","") 
  Response.Write "<center>请关闭窗口!</center>"
  Response.End()
  
End If

'Session("MemID")  = MemID
'<input type='checkbox' name='PrmList' value='xMemFEditor'>(编辑器)附件上传　　
'<input type='checkbox' name='PrmList' value='xMemFUpload'>(附图)附件上传　
' Session("MemPerm") = "{("&MemID&":"&MemGrade&"("&MemType&");}"&gMemPerm(conn,MemGrade) 

%>
<%End If%>
</body>
</html>
