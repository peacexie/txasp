
<!--#include file="../../sadm/func2/func_const.asp"-->
<!--#include file="../../sadm/func2/func_perm.asp"-->
<!--#include file="../../sadm/func1/func_time.asp"-->
<!--#include file="../../sadm/func1/func_vbs.asp"-->
<!--#include file="../../sadm/func1/func_rs.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->

<%

SysPassPW = "my_ext_pass"&Config_Code ' 单独登陆工具密码 
SysDBType = cfgDBType  ' Access,MsSql
'Response.Write SysMod&SysDBType
conDB = conn 

%>
