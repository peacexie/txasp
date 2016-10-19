
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->

<!--#include file="../../sadm/func2/func_sfile.asp" -->
<!--#include file="../../upfile/sys/para/tmstyp2.asp" -->
<!--#include file="../../upfile/sys/para/tmsPara.asp" -->

<!--#include file="../../upfile/sys/config/_depart.asp"-->

<!--#include file="../file/cfg-mpath.asp"-->

<%

'Response.Write PrmID
PrmFlag = Request("PrmFlag")
Call Chk_Perm9(ModID,PrmFlag)

upRoot = Config_Path&"upfile/"

%>

