<!--#include file="../himg/tconfig.asp"-->
<%
For i = 0 To 10
  x = Int((65536 - 32768) * Rnd + 32768)'Get_YYYYMMDD("")&"-"&Get_HHMMSS&"-"&Rnd_ID("KEY",3) 
  Response.Write "<a href='?test="&x&"'>test_"&x&"</a> <br>"
Next
%>
Path:<%=Config_Path%><%=Server.URLEncode("#")%>
<a 
href="Online.asp?send=View">View</a> - <a 
href="Online.asp?send=Clear">Clear</a> - <a 
href="Online.asp?FlgMod=Home">Home</a>
<script src="Online.asp"></script>
