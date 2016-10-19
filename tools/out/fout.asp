<!--#include file="../../sadm/func1/func_obj.asp"-->
<!--#include file="../../sadm/func1/upremote.asp"-->
<%

'// Functions 
Function getRate()
  Dim Url,sHtml,sPos,nRate
  Url = "http://open.baidu.com/huilv/s?wd=1%C3%C0%D4%AA%B5%C8%D3%DA%B6%E0%C9%D9%C8%CB%C3%F1%B1%D2%D4%AA&tn=baiduhuilv"
  sHtml = OutPage(Url,"gb2312")&"" 
  sPos = inStr(sHtml,"num2 = numFormat")
  nRate = Mid(sHtml,sPos,32)
  nRate = OutSFlag(nRate,"(",")")
  getRate = nRate
End Function

Function getFee(c,m,w,xRate)
  Dim Url,sHtml,sTFee,sTime,sTrack,feeInfo,res(1),w2,r2
  w = Replace(w,",","")
  w2 = w'w*1.2
  Url = "http://www.sendfromchina.com/shipfee/out_rates/?country="&c&"&mode="&m&"&weight="&w2&""
  sHtml = OutPage(Url,"iso-8859-1")&""  
  If inStr(sHtml,"</totalfee>")>0 Then
    sTFee = OutSFlag(sHtml,"<totalfee unit=""RMB"">","</totalfee>")
    sTime = OutSFlag(sHtml,"<deliverytime unit=""WorkDay"">","</deliverytime>")
    sTrack = OutSFlag(sHtml,"<iftracking>","</iftracking>")
  End If
  sTFee = Replace(sTFee,",","")
  If isNumeric(sTFee) Then
	If sTFee<0 Then
	  sTFee = 0   
	Else
	  If inStr(",HKBRAM,SGRAM,CNRAM,",m&",")>0 Then
	    r2 = 1.35 '1.5
	  ElseIf inStr(",CNSFEDEX,",m&",")>0 Then 'CNAP,CNSFEDEX,WWEMS
	    r2 = 1.25 ' 1.30, 1.25, 1.20, 1.15
	  Else
	    r2 = 1.20 ' 1.15
	  End If
	  sTFee = sTFee*r2/xRate 'sTFee/xRate
	End If
  Else
	  sTFee = 0 
  End If
  res(0) = FormatNumber(sTFee,2)
  res(1) = "$"&res(0)&" &nbsp; WorkDay:"&sTime&" &nbsp; Tracking:"&sTrack&""
  getFee = res
End Function


%>



