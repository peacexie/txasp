<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/admin/sys.funcs.asp"-->
<%

If Request("ModID")<>"" Then
  ModID = RequestS("ModID","C",48)
  Session("ModID") = ModID
ElseIf Session("ModID")&""<>"" Then
  ModID = Session("ModID")
Else
  ModID = "InfN124"
End If


If Session("UsrID")&""="" And Session("MemID")="guest" Then 
  Response.End()
End If
'admin//sys_dbacc.asp,sys_config.asp,sys_menu.asp,sys_mpara.asp,type_list.asp,typs_list.asp
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) '/sadm/system/s-config.asp
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1) 's-config.asp
If PrmFile="type_list.asp" Then
  PrmID = ModID 
ElseIf PrmFile="type_page.asp" Then
  PrmID = ModID
'ElseIf PrmFile="type_center.asp" Then
  'PrmID = ModID
Else
  PrmID = "" 
End If
'Response.Write PrmID
Call Chk_Perm1(PrmID,"")
'Call Chk_URL()


If inStr(ModID,"Pic")>0 Then
 PathImg = Config_Path&"upfile/pics/"
 TypeFile = ""
 TypeImg = ".JPG/.GIF/.JPEG"
 ModTab = "InfoPics"
ElseIf inStr(ModID,"InfV")>0 Then
 PathImg = Config_Path&"upfile/pics/"
 TypeFile = ""
 TypeImg = ".JPG/.GIF/.JPEG"
 ModTab = "InfoPics"
'ElseIf inStr(ModID,"InfP")>0 Then
 'PathImg = "../../upfile/pics/"
 'TypeFile = ""
 'TypeImg = ".JPG/.GIF/.JPEG"
 'ModTab = "InfoPics"
Else
 PathImg = Config_Path&"upfile/news/"
 TypeFile = ".DOC/.XLS/.TXT/.HTML/.HTM"
 TypeImg = ".JPG/.GIF/.JPEG"
 ModTab = "InfoNews"
End If


%>

