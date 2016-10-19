<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<%

Dim edcRootUrl,edcEditID,edcEditUrl
Dim flgScr,flgPara
flgScr = "script" :flgPara = "type='text/javascript' charset='utf-8'"

edTestMod = Request("edTestMod")&""
edTBakID = edNameID '备份
If edTestMod<>"" Then
  edTempNM = Request("edTempNM") 
  edNameID = edTestMod
Else
  edTempNM = "(Defalut)Editor"
End If


edcRootUrl = Config_Path 
edcRootUrl = Left(Config_Path,Len(Config_Path)-1) ' /php
edcEditID = edNameID ':edcEditID = "edxh1" '//edck3,edxh1,edkind,edfck,
edcEditUrl = edcRootUrl&"/sadm/"&edcEditID 
edcFileExt = lCase(Request.ServerVariables("SCRIPT_NAME"))
edcFileExt = Mid(edcFileExt,InStrRev(edcFileExt,"."))


'  'bmp'
imageExtStr = "gif,jpg,jpeg,png"
flashExtStr = "swf,flv"
mediaExtStr = "swf,flv,mp3,wav,wma,wmv,mid,avi,mpg,asf,rm,rmvb"
fileExtStr = "doc,docx,xls,xlsx,ppt,htm,html,txt,zip,rar,gz,bz2"


Function Get_Para(xID,xDef)
	Dim re :re = Request(xID)&""
	If re="" Then re = xDef
	Get_Para = re
End Function

' 判断权限，for Editor
Function Chk_PEditor(xSyst)
'// FileInfo, 信息附件上传'// FileEditor, 编辑器上传'// FileView, 服务器浏览'// FileAdmin, 服务器管理上传'// Chk_Perm9("MemFEditor","3");	
	Dim re,sPerm
	If xSyst="" Then xSyst = "FileEditor"
	sPerm = Session(UsrPStr)&Session("MemPerm")&Session("InnPerm")&""
	if(sPerm="") Then
	  re = "NoLogin"
	ElseIf (inStr("("&sPerm&")","{Admin}")>0) Then
	  re = "(Pass)"
	ElseIf (inStr(sPerm,xSyst)>0) Then
	  re = "(Pass)"
	Else
	  re = xSyst
	End If
	Chk_PEditor = re
End Function

Function json_Alert(xFlag,xValue,xType)
	Dim hash
	If xType="" Then xType = "msg"
	Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
	Set hash = jsObject()
	If xType="msg" Then
	  hash("error") = xFlag
	  hash("message") = xValue
	Else
	  hash("error") = xFlag
	  hash("url") = xValue
	End If
	hash.Flush
	Set hash = Nothing
	Response.End
End Function

%>