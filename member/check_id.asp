<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../sadm/func1/func1.asp"-->
<!--#include file="../sadm/func2/func2.asp"-->
<!--#include file="../upfile/sys/para/rkeyid.asp"-->
<!--#include file="../inc/form/form_app.asp"-->

<%

  MemID = LCase(RequestS("MemID","C",64))
  verMemb = Request("verMemb")
  yAct = Request("yAct")
  ChkCode = uCase(Request("ChkCode"))
  
  
If verMemb="2" Then

  vMem_CIDA1 = " Please input a Correct ID !"
  vMem_CIDA2 = " Sorry, the user ID has used! "
  vMem_CIDA3 = " Error: user ID has some keywords:<br>"
  vMem_CIDA4 = "Sorry, the user ID has used! <br>Please change one."
  vMem_CIDA5 = "OK, this user ID can be used by your! "
  
  vMem_CIDB1 = "Check User ID"
  vMem_CIDB2 = "The keywords as below is not for apply:"
Else

  vMem_CIDA1 = " 请正确输入 !"
  vMem_CIDA2 = " 对不起，帐号 已经存在 或 已经有其它用途! "
  vMem_CIDA3 = " 错误：帐号含有以下关键字:<br>"
  vMem_CIDA4 = "对不起,帐号已经被注册,<br>请选用其他帐号."
  vMem_CIDA5 = "恭喜您,您可以注册此帐号."
  
  vMem_CIDB1 = "帐号检查"
  vMem_CIDB2 = "以下关键字不能申请:"
  
End If
  
  
  sql1 = "SELECT MemID FROM [Member"&Mem_aMemb&"]  WHERE MemID='"&MemID&"'"
  sql2 = "SELECT UsrID FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&MemID&"'"
  if Len(MemID&"") < 4 then
    msg = vMem_CIDA1
  ElseIf rs_Exist(conn,sql2)="YES" then
    msg = vMem_CIDA2
  elseif inStr(rKeyID&",",MemID&",")>0 then
   msg = vMem_CIDA3&rKeyID
  elseif rs_Exist(conn,sql1) = "YES" then
    msg = "[<font color=blue>"&MemID&"</font>] <br><font color=red>"&vMem_CIDA4&"</font>"
  else
    msg = "[<font color=blue>"&MemID&"</font>] <br><font color=3333FF>"&vMem_CIDA5&"</font>"
  end if
  
''//////////////////////// 
'' ChkAjID,ChkAjCode
'' var url = "check_id.asp?yAct="+xType+"&MemID="+fmMemID.value+"&ChkCode="+fmMemID.ChkCode; 

If yAct = "ChkAjCode" Then
  rFlag = "Y.Code"
  If Session("ChkCode")<>ChkCode Then
    rFlag = "N.Code" 
  End If
End If
If yAct = "ChkAjID" Then
  rFlag = "Y.ID"
  If Len(MemID&"") < 4 Then
    rFlag = "N.<4"
  ElseIf rs_Exist(conn,sql2)="YES" then
    rFlag = "N.Adm"
  Elseif inStr(rKeyID&",",MemID&",")>0 And Session("UsrID")&""="" then
    rFlag = "N.Key"
  Elseif rs_Exist(conn,sql1) = "YES" then
    rFlag = "N.Mem"
  End If
End If

If yAct="ChkAjID" OR yAct="ChkAjCode" Then
  Call Chk_Url()
  t1=Show_Form(Request.Servervariables("HTTP_REFERER"))
  t2=Show_Form(Request.ServerVariables("HTTP_HOST"))
  Response.Write rFlag
  'Response.Write " - "&t
  Response.End()
ElseIf yAct="ChkAjPW" Then '//2011-04-21:未使用
  Call Chk_Url()
  App26Str = ""
  App26Day = Get_AppDay()+47 '每天变化一次
  For i=1 To Len(App26Code)
    c = Mid(App26Code,i,1)
    If c=Chr(34) Then
	  c="'"&c&"'"
	Else
	  c=""""&c&""""
	End If
	App26Str = App26Str&"<input name="&c&" type='hidden' value='"&Chr(App26Day+i)&"' />"&vbcrlf
  Next
  Response.Write App26Str 'Response.Write App26Code 
  Response.End()
ElseIf yAct="ChkAj22" Then
  Call Chk_Url()
  Session(App22Code)=App22Data
  Response.Write App22Data
  Response.End()
End If
''//////////////////////// 
  
  rKeyID = Replace(rKeyID,",",", ")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%="[ "&Config_Name&" ]"%> - <%=vMem_CIDB1%></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
<!--
body, td, th {
	font-size: 13px;
	line-height: 150%;
}
-->
</style>
</head>
<body>
<table width="360" height="100%"  border="0" align="center" cellpadding="0" cellspacing="0" style="word-wrap:break-all;">
  <tr>
    <td align="center"><%=msg%></td>
  </tr>
  <tr>
    <td align="center" valign="bottom"><%=vMem_CIDB2%><br>
      <%=rKeyID%></td>
  </tr>
</table>
</body>
</html>