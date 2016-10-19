<!--#include file="../tools/out/err404.asp"-->
<!--#include file="func2/func_const.asp"-->

<%
Dim act : act=lCase(Request("act"))
If act="admin" Or act="login" Then
  Response.Write "<!--"&Config_RAdm&".asp"&"-->"
  'Response.Redirect Config_RAdm&".asp"
End If
%>