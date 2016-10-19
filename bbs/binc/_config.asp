<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../badm/conpub.asp" -->

<!--#include file="../../sadm/func2/func_sfile.asp" -->


<%


If SwhBBSStop="Y" Then 
  Response.Write File_Read("../upfile/sys/para/bbsstop.htm","utf-8") 
  Response.End()
End If


'说明                                            ' 公开论坛 | 内部论坛
'其中 bbsPath项目                                ' 如果配置公开论坛,内部论坛两个论坛，则要两个目录，否则只用这个目录
'_itop.asp文件中的会员连接项，根据需要修改; 
'bbs_list.asp论坛帖子管理的查看连接，可做稍微修改...

' 配置开,内部两个论坛 

ModID = "BBSP124"                                ' BBSP124 | BBSI124
ModTab = "BBSInfo"                               ' 公共 有必要的话，用表BBSInf2 .... 分开内部帖子
upRoot = Config_Path&"upfile/"                      ' 公共 
  
bbsName = "(论坛名称)"                           ' (论坛名称) | (xx公司 内部论坛)
bbsPath = Config_Path&"bbs/"                        ' Config_Path"/bbs/" | Config_Path"/bis/" 
bbsUser = Session("MemID")                       ' Session("MemID") | Session("InnID")
bbsLogin = Config_Path&"member/login.asp"           ' Config_Path&"member/login.asp" | Config_Path&"ext/login.asp
bbsUPass = Config_Path&"inc/mem_inc/mem_main.asp"   ' Config_Path&"inc/mem_inc/mem_main.asp" | Config_Path&"doc/userpw.asp 

If ModID = "BBSP124" Then
  '// 发贴检查权限？！// 用文件名判断，这里未作判断 
  'Call Chk_Perm2("xBBSInfo",bbsPath) '任何时候，登陆进入
ElseIf ModID = "BBSI124" Then
  '任何时刻，检查权限
  Call Chk_Perm3(ModID,bbsPath)
Else
  '严格来讲，是非法的
  Call Chk_Perm2("xBBSInfo",bbsPath)
  Call Chk_Perm3(ModID,bbsPath)
End If

Set rs=Server.CreateObject("Adodb.Recordset")

%>