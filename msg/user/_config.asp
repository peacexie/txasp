<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../config/config.asp"-->
<%

PrmUsr = "ModSms"
PrmMem = "xTradeSms" '商务短信发送

PrmFlag = Request("PrmFlag")
If PrmFlag="(Mem)" Then
  PrmUser = Session("MemID")
Else
  PrmUser = Session("UsrID")
End IF
'user//index.asp,user_editpw.asp,...
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) 
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1)
If PrmFile="sapi.asp" Then 
    'Pass
ElseIf inStr("member.asp;madd.asp;medit.asp",PrmFile)>0 Then
    Call Chk_Perm1(PrmUsr,"")
ElseIf inStr("tels.asp;temp.asp;type.asp;sign.asp;tset.asp",PrmFile)>0 Then
  If PrmFlag="(Mem)" Then
    Call Chk_Perm2(PrmMem,"")
  Else
    Call Chk_Perm1(PrmUsr,"")
  End IF 
Else
    Call Chk_Perm2(PrmMem,"")
End If

%>