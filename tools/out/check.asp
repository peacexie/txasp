<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<%
Set rs=Server.CreateObject("Adodb.Recordset")
act = Request("act") : If act="" Then act="[def]" 
'Advert,Update,Index,Info,Vers,TestAdv,TestUpd,---chkTime,
scr = "script"

If inStr("[Advert,Update,TestAdv,TestUpd]",act)>0 Then
  
 '/// Update更新 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 If act="Update" Then
  '检查
  Response.Write "document.write('Update更新');"
 '/// Advert广告 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 ElseIf act="Advert" Then
  Response.Write "document.write('Advert广告');" 
 '/// TestAdv广告 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 ElseIf act="TestAdv" Then
  Response.Write "<"&scr&" src='?act=Advert' type='text/java"&scr&"'></"&scr&">"
 '/// TestUpd广告 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 ElseIf act="TestUpd" Then
  Response.Write "<"&scr&" src='?act=Update' type='text/java"&scr&"'></"&scr&">" 
 End If

Else 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="robots" content="noindex, nofollow">
<meta http-equiv="refresh" content="180">
<style type="text/css">
body, td, th { font-size: 12px; }
body { margin: 1px; }
</style>
</head>
<body>
<%
 '/// Info系统信息 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 If act="Index" Then
  Response.Write "<a href='?act=Info'>Info</a> - <a href='?act=Vers'>Vers</a> - <a href='?act=Advert'>Advert</a> - <a href='?act=Update'>Update</a> - <a href='?act=[def]'>[Default]</a> - <a href='?act=TestUpd'>TestUpd</a> - <a href='?act=TestAdv'>TestAdv</a>"
 ElseIf act="Info" Then
  sql = "SELECT * FROM [xqVers] Order BY ParCode " 
  rs.Open sql,conn,1,1 
  Do While NOT rs.EOF
    sID = rs(0)
    sName = rs(1)
    sExt = rs(2)
    Response.Write sID&" ----- "&sName&" ----- "&sExt&"<BR>"
  rs.MoveNext
  Loop
  rs.Close()
 '/// Vers当前版本 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 ElseIf act="Vers" Then
  Response.Write rs_Val("","SELECT ParName FROM AdmPara WHERE ParCode='dPlnVerNow'")
 
 '/// ..... XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 'ElseIf act="...." Then
 ' Response.Write ....
  
 '/// Session检查 XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 Else
  If ""&Session("UsrID")&Session("MemID")&Session("InnID")="" Then  
    sFlag = "NG!" 
  Else
    sFlag = "OK!" 
  End If
  sFlag = sFlag&" "&Timer()
  Response.Write "<font color='#0000FF'>"&sFlag&"</font>"
 '/// End XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 End If

%>
<%
  If Request("rem")="OK" Then
    Response.Write vbcrlf&"<"&scr&" type='text/javascript'>"
%>
<!--#include file="../../upfile/sys/para/remind.js"-->
<!--#include file="chkremind.js"-->
<%
    Response.Write vbcrlf&"</"&scr&">"
  End If
%>

</body>
</html>
<%
End If
Set rs=Nothing
%>

