
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp" -->

<!--#include file="../../upfile/sys/config/_inner.asp"-->
<!--#include file="../../upfile/sys/para/tmstyp2.asp" -->

<%

sysName = "(xx公司 内部公文)"
ModID = "DocD124"
ModTab = "DocsNews"
upRoot = Config_Path&"upfile/"
upPart = "dtdoc"
Session("upPart") = upPart

Call Chk_Perm1(ModID,"")

%>
