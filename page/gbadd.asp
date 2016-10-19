<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="_config.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%
MD = "GboK124"
ModTab = rel_ModTab(MD)
upPart = rel_TabPath(ModTab)
upRoot = Config_Path&"upfile/"
'Response.Write "11:"&ModTab&upPart

TP = RequestS("TP",3,255)
ID = RequestS("ID",3,255)

InfType = RequestS("InfType",3,24)
ChkCode = Request("ChkCode")

  Dim sys27_Rnd(10)
  sys27_RVal = Request(App27Random)&""
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If

send = Request("send")
If send="ins" Then 
  Call Chk_Url()   'KeyID = Get_AutoID(24)
  KeyID = rs_AutID(conn,"GboInfo","KeyID",upPart,"1","")
InfSubj = RequestS("InfSubj"&sys27_Rnd(1),C,255)
InfCont = RequestS("InfCont"&sys27_Rnd(2),C,24000) 
'// InfReply = vGbo_RepNO2

'// 判断过滤
fSubj = Chr_Fil2(InfSubj)
fCont = Chr_Fil2(InfCont)
  SetShow = "Y"
  fMsgShow = vGbo_SendOK1
If SwhGbkChk="N" Then 
  SetShow = "N"
  fMsgShow = vGbo_SendOK2
ElseIf fSubj<>InfSubj or fCont<>InfCont then
  SetShow = "N"
  fMsgShow = vGbo_SendOK3
  InfSubj = fSubj
  InfCont = fCont
Else

End If

'// 检查编辑器
If SwhGbkEditor="Y" Then
  InfCont = Show_Html(InfCont)
Else
  InfCont = Show_Text(InfCont)
End If
  
'// 检查是否存文件
If Config_Cont="DB" Then
  xxxCont = InfCont
  xxxReply = ""'InfReply
Else
  xxxCont = ""
  xxxReply = ""
End If

'fMsgShow = fMsgShow&ParFilALLKeys

sql = " INSERT INTO [GboInfo] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyFlag" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", LnkName" 
sql = sql& ", LnkFace" 
sql = sql& ", LnkEmail" 
sql = sql& ", SetRead" 
sql = sql& ", SetShow" 
sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ", LogEditIP" 
sql = sql& ", LogEUser" 
sql = sql& ", LogETime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '"&MD&"'" 
sql = sql& ", '" & RequestS("TreeTypA",C,48)&"^"&RequestS("TreeTypB",C,48) &"'" 
sql = sql& ", '" & RequestS("InfType",C,255) &"'"
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & RequestS("LnkName"&sys27_Rnd(4),C,24) &"'" 
sql = sql& ", '" & RequestS("LnkFace",C,48) &"'" 
sql = sql& ", '" & RequestS("LnkEmail"&sys27_Rnd(5),C,255) &"'" 
sql = sql& ", 0" 
sql = sql& ", '" & SetShow &"'" 
sql = sql& ", '" & RequestS("ImgName",C,255) &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("MemID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '-'" 
sql = sql& ", '-'" 
sql = sql& ", '1900-12-31'" 
sql = sql& ")"

If  Session("ChkCode")<>uCase(ChkCode) Then
  fMsgShow = vMsg_ChkCErr
ElseIf Len(InfSubj)>0 And Len(InfCont)>0 Then
  Call rs_DoSql(conn,sql) 
  upPath = upRoot&Replace(KeyID,"-","/") 
  Call add_sfFile()
Else
  fMsgShow = vGbo_SendNG
End If

  Response.Write js_Alert(fMsgShow,"Redir","gbook.asp"&"?ModID="&MD&"") 

End If

%>


