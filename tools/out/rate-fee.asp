<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
<%

'// Paras
Dim c,m,w,id
Dim Url,sHtml,sPos
Dim sRate,sTFee,sTime,sTrack

c = Request("c")
m = Request("m")
w = Request("w")
id = Request("id")
If id="" Then
 If c="" Then c="Canada"
 If m="" Then m="CNUPS"
 If w="" Then w="10"
End If


'// Functions 
Sub getRate()
  Url = "http://open.baidu.com/huilv/s?wd=1%C3%C0%D4%AA%B5%C8%D3%DA%B6%E0%C9%D9%C8%CB%C3%F1%B1%D2%D4%AA&tn=baiduhuilv"
  sHtml = OutPage(Url,"gb2312")&"" 
  sPos = inStr(sHtml,"num2 = numFormat")
  sRate = Mid(sHtml,sPos,32)
  sRate = OutSFlag(sRate,"(",")")
End Sub

Sub getFee()
  Url = "http://www.sendfromchina.com/shipfee/out_rates/?country="&c&"&mode="&m&"&weight="&w&""
  sHtml = OutPage(Url,"iso-8859-1")&""  
  sTFee = OutSFlag(sHtml,"<totalfee unit=""RMB"">","</totalfee>")
  sTime = OutSFlag(sHtml,"<deliverytime unit=""WorkDay"">","</deliverytime>")
  sTrack = OutSFlag(sHtml,"<iftracking>","</iftracking>")
End Sub


'// get Paras
If id="" Then
  Call getFee()
ElseIf id="Rate" Then
  Call getRate()
ElseIf id="Fee" Then
  Call getFee()
Else
  Call getRate()
  Call getFee()
End If


'// formate Paras
If isNumeric(sTFee) And isNumeric(sRate) Then
If sTFee>0 And sRate>0 Then
  sTFee = FormatNumber(sTFee/sRate,2) 
End If
End If
feeInfo = "$"&sTFee&" &nbsp; WorkDay:"&sTime&" &nbsp; Tracking:"&sTrack&""
'feeInfo = feeInfo&"<input name=feeDiv type=hidden id=feeDiv value=\'"&sTFee&"\' >"
'feeInfo = feeInfo&"("&sRate&")" '/// b8


'// Show Paras
If id="" Then
  Response.Write vbcrlf&feeInfo
  Response.Write vbcrlf&"<br>?c=Australia&m=HKUPS&w=2.4&id="
  Response.Write vbcrlf&"<br>?c=Australia&m=HKUPS&w=2.4&id=3"
  Response.Write vbcrlf&"<br>?id=Rate"
  Response.Write vbcrlf&"<br>?c=Australia&m=HKUPS&w=2.4&id=Fee"
ElseIf id="Rate" Then
  Response.Write vbcrlf&sRate
ElseIf id="Fee" Then
  Response.Write vbcrlf&feeInfo
Else 'If id<>"" Then
  If(Len(sTFee)=0) Then
	Response.Write vbcrlf&"alert('Error: "&id&", Select a nother one!');"
	Response.Write vbcrlf&"getElmID('rad_'+"&id&").checked = false;"  
  Else
	Response.Write vbcrlf&"getElmID('div_'+"&id&").innerHTML = '"&feeInfo&"';"  
  End If
  Response.Write vbcrlf&"getElmID('bSum').innerHTML = '"&sTFee&"';" 
  Response.Write vbcrlf&"cal_cFee();"
End If


'// for Test
'Response.Write vbcrlf&"alert('"&id&"');"

%>



