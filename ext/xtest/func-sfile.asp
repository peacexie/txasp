<!--#include file="../../sadm/func2/func_sfile.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../upfile/sys/para/tmspara.asp"-->
<!--#include file="../../upfile/sys/para/tmstyp2.asp"-->

<%



sTmpID = get_TmpID("Vote","xInfN888;xN810012;") 
Response.Write "<br>"&sTmpID

sTmpID = get_TmpID("List","xInfN888;xN810012;") 
Response.Write "<br>"&sTmpID
' InfN888;N810012;

sPath = Split("dtdef-2008-","-")(0)
Response.Write "<br>"&sPath

%>


