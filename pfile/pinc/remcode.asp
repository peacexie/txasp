<!--#include file="../../upfile/sys/para/keywords.asp"-->
<%

If verNow="2" Then
 vIRem_Search = "Search"
 vIRem_SchMsg = "Search"
 vIRem_SchKey = "Keyword"
 vIRem_OrgPage = "Original"
 vIRem_Add = "Write Remark"
 vIRem_Subj = "Subject"
 vIRem_Cont = "Content"
 vIRem_Name = "Name"
 vIRem_BtnOK = "OK Send"
 vIRem_BtnNO = "Reset"
 vIRem_Attitude = "Attitude"
 vIRem_AttOK = "Agree"
 vIRem_AttNG = "Oppose"
 vIRem_AttMid = "Neutrally"
 vIRem_MsgCode = "Check Code Error,Please Click the Code Picture And Send Again! "
 vIRem_MsgNull = "Send Error, Subject or Content Can NOT Null!"
 vIRem_MsgOK = "Send OK!"
 vIRem_MsgKeys = "Send OK!! \n But the Massage will be Check by the Web Master! "
 vIRem_MsgNG = "Send Error!"
Else
 vIRem_Search = "搜索"
 vIRem_SchMsg = "评论搜索"
 vIRem_SchKey = "关键字"
 vIRem_OrgPage = "原文"
 vIRem_Add = "添加评论"
 vIRem_Subj = "主题"
 vIRem_Cont = "内容"
 vIRem_Name = "姓名"
 vIRem_BtnOK = "提交信息"
 vIRem_BtnNO = "重新填写"
 vIRem_Attitude = "态度"
 vIRem_AttOK = "支持"
 vIRem_AttNG = "反对"
 vIRem_AttMid = "观望"
 vIRem_MsgCode = "认证码错误！请点击认证码图片，刷新认证码再提交一次"
 vIRem_MsgNull = "空内容禁止提交!"
 vIRem_MsgOK = "提交评论成功!"
 vIRem_MsgKeys = "感谢您的评论! \n但信息需要审核! "
 vIRem_MsgNG = "提交评论失败!"
End If


ModTab = "GboSend"
upPart = "defdt"
upRoot = Config_Path&"upfile/"

send = RequestS("send","C",12)
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,48)


Dim sys27_Rnd(10)
If send = "send" Then
  sys27_RVal = Request(App27Random)&""
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If
Else
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next
End If


ObjSubj = RequestS("ObjSubj","C",255) 
ObjID = RequestS("ObjID","C",48)
ObjUrl = RequestS("ObjUrl","C",255) 
If ObjUrl="" Then
ObjUrl = Request.Servervariables("HTTP_REFERER")
End If
'Response.Write ObjSubj&ObjID&ObjUrl
 
send = Request("send")
ChkCode = uCase(Request("ChkCode"))
'Response.Write ChkCode&"dd"

If send="send" Then

  Call Chk_URL()	
  'Response.End()
  IP = Get_CIP()
  InfSubj = RequestS("InfSubj"&sys27_Rnd(1),3,255)
  InfCont = RequestS("InfCont"&sys27_Rnd(2),3,2400)
  LogAUser = RequestS("LogAUser"&sys27_Rnd(4),3,48)
  InfCont = Show_Text(InfCont)
  
  fName = Chr_Fil2(LogAUser)
  fSubj = Chr_Fil2(InfSubj)
  fCont = Chr_Fil2(InfCont)
    RemShow = "N"
    If SwhRemChk="Y" Then
        RemShow = "Y"
	  if fName<>LogAUser or fSubj<>InfSubj or fCont<>InfCont then
        RemShow = "N" 
        LogAUser = fName 
        InfSubj = fSubj 
        InfCont = fCont
      end if
    End If
  fADD = "Pass" ': Response.Write ChkCode
  If Session("ChkCode")<>ChkCode Or Session("ChkCode")&""="" Then
    fMsgShow = vIRem_MsgCode&Session("ChkCode")&":"&ChkCode
	PicReLoad = "PicReLoad('../');" 
	Response.Write js_Alert(fMsgShow,"Alert","-1")
	Response.End()
  ElseIf Len(InfSubj)=0 Or Len(InfCont)=0 Then
    fADD = vIRem_MsgNull
	Response.End()
  End If
  
  InfCont = fCont '还原
  If Config_Cont="DB" Then
    xxxCont = InfCont
  Else
    xxxCont = ""
  End If
  
  KeyID = rs_AutID(conn,"GboInfo","KeyID",upPart,"1","")
sql = " INSERT INTO [GboSend] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont"  
sql = sql& ", SetShow"  
sql = sql& ", LogAddIP" 
sql = sql& ", LogATime" 
sql = sql& ", LogAUser" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & MD & "'" 
sql = sql& ", '" & ObjID &"'" 
sql = sql& ", '" & RequestS("InfType","C",24) &"'" 
sql = sql& ", '" & fSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & RemShow &"'" 
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '" & RequestS("LogAUser"&sys27_Rnd(4),"C",48) &"'" 
sql = sql& ")"

  If fADD="Pass" Then 
    Call rs_DoSql(conn,sql)  
    upPath = upRoot&Replace(KeyID,"-","/") 
    Call add_sfFile()
    rMsg = vIRem_MsgOK
    If RemShow="N" Then
      rMsg = vIRem_MsgKeys
    End If
	Response.Write js_Alert(rMsg,"Redir","?ModID="&MD&"&ObjID="&ObjID&"&ObjSubj="&Server.URLEncode(ObjSubj)&"&ObjUrl="&ObjUrl&"")
  Else
      rMsg = vIRem_MsgNG&fADD
  End If ' End Pass
  
End If

ObjSubj = Show_Text(ObjSubj)
 
%>