<!--#include file="_config.asp"-->
<%
  s=VotList("BBSVB24",1,"Now",120)
If s="" Then
  s=VotList("BBSVB24",1,"All",120)
End If

If s<>"" Then
  p=inStr(s,"?ID=")
  ID=Mid(s,p+4,24)
Else
  ID=""
End If

If ID<>"" Then
  Response.Redirect "vote.asp?ID="&ID
Else
  Response.Redirect "vlist.asp?TimID=All"
End If

'Response.Write Server.HTMLEncode(ID&"<br>"&s)
'Response.Redirect "vlist.asp"
%>