<!--#include file="../func1/func1.asp"-->
<!--#include file="../func2/func2.asp"-->
<!--#include file="../func1/func_file.asp"-->
<!--#include file="../func1/func_opt.asp"-->
<!--#include file="sys.funcs.asp"-->
<%

'admin//sys_dbacc.asp,sys_config.asp,sys_menu.asp,sys_mpara.asp,type_list.asp,typs_list.asp
PrmPath = LCase(Request.ServerVariables("PATH_INFO")) '/sadm/system/s-config.asp
PrmFile = Mid(PrmPath,InStrRev(PrmPath,"/")+1) 's-config.asp
If PrmFile="sys_dbacc.asp" Or PrmFile="sys_dbinfo.asp" Or PrmFile="typs_list.asp" Then
  PrmID = "SysDB" 
ElseIf PrmFile="sys_config.asp" Or PrmFile="type_list.asp" Then
  PrmID = "SysConfig" 
ElseIf PrmFile="sys_menu.asp" Or PrmFile="sys_mpara.asp" Then
  PrmID = "SysMenu" 
Else
  PrmID = "ModSystem" '"System"
End If

'Response.Write PrmID
Call Chk_Perm1(PrmID,"")

%>

