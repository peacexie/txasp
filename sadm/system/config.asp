<!--#include file="../func1/func1.asp"-->
<!--#include file="../func2/func2.asp"-->
<!--#include file="../func1/func_file.asp"-->
<!--#include file="../func1/func_opt.asp"-->
<%

'system//logs_list.asp,system.asp,para_xxx
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) '/sadm/system/s-config.asp
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1) 's-config.asp
If PrmFile="logs_list.asp" Or inStr(PrmFile,"para_")>0 Or PrmFile="upd_para.asp" Then
  PrmID = "SysMods" 
ElseIf inStr("system.asp",PrmFile)>0 Then
  PrmID = "SysMods" 
  If Request("ModID")="Inner" Then
    PrmID = "ModDocs" '内部公文 
  End If
Else
  PrmID = "ModSystem" '"System"
End If

'Response.Write PrmID
Call Chk_Perm1(PrmID,"")

%>