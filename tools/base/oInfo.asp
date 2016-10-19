<!--#include file="../himg/tconfig.asp"-->
<!--#include file="oConfig.asp"-->
<%
' OnlMax: /web/OnlAct.asp中更新
'If Application("OnlMax")&""="" Then
  
'End If
Application("OnlMax") = rs_Val("","SELECT vNum FROM OnlView WHERE vID='OnlMax'")
Application("OnlNow") = rs_Count(conn,"OnlPage WHERE pTime>="&cfgTimeC&""&DateAdd("n",(-1)*pExpTime,Now())&""&cfgTimeC&"")
AllView = rs_Val("","SELECT vNum FROM OnlView WHERE vID='AllView'")

Response.Write vbcrlf&"PeaceOnlNow.innerHTML='"&Application("OnlNow")&" 人';"
Response.Write vbcrlf&"PeaceOnlMax.innerHTML='"&Application("OnlMax")&" 人';"
Response.Write vbcrlf&"PeaceAllView.innerHTML='"&AllView&" 次';"

' /web/OnlView.asp 
If Request("send")="View01" Then
  OnlMTime = rs_Val("","SELECT vTime FROM OnlView WHERE vID='OnlMax'")
  DatePrev = rs_Val("","SELECT vNum FROM OnlView WHERE vID='"&DateAdd("d",-1,Date())&"'")
  DateThis = rs_Val("","SELECT vNum FROM OnlView WHERE vID='"&Date()&"'")
  Response.Write vbcrlf&"PeaceOnlMTime.innerHTML='"&OnlMTime&"';"
  Response.Write vbcrlf&"PeaceDatePrev.innerHTML='"&DatePrev&" 次';"
  Response.Write vbcrlf&"PeaceDateThis.innerHTML='"&DateThis&" 次';"
End If

%>