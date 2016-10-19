<!--#include file="../func1/func1.asp"-->
<!--#include file="../func2/func2.asp"-->
<!--#include file="../func1/func_file.asp"-->
<!--#include file="../func1/func_opt.asp"-->
<%

'user//index.asp,user_editpw.asp,...
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) 
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1) 
If inStr("index.asp;user_editpw.asp",PrmFile)>0 Then
  PrmID = "" 
Else
  PrmID = "SysUList" '"System"
  If Request("ModID")="Inner" Then
    PrmID = "ModDocs" '内部公文 
  End If
End If

'Response.Write PrmID
If PrmFile<>"user_editpw.asp" Then
  Call Chk_Perm1(PrmID,"")
End If

%>