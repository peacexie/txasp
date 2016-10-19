<!--#include file="inc_func.asp"-->
<%
If Session("PeaceDB")&""<>PassID Then
  Response.Redirect "index.asp"
  Response.End()
Else
  Response.Write("<a href='index.asp'>Home</a> | <a href='index.asp?yAct=Out'>Logout</a>")
End If
%>

