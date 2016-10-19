<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp" -->
<%

'News, Jobs, Pics, Vdos
'RssID,Title,TabID,ModID,TypID,ListUrl,ViewUrl
sRssID = "InfN124|PicS124"
sTitle = "新闻中心|产品图片"
sTabID = "InfoNews|InfoPicS" 
sModID = "InfN124|PicS124"
sTypID = "|" 'N110048
sUrlLI = "../page/info.asp?ModID=(ModID)&TypID=(TypID)|../page/info.asp?ModID=(ModID)"
sUrlVW = "../page/iview.asp?KeyID=(KeyID)|../page/iview.asp?KeyID=(KeyID)"

aRssID = Split(sRssID,"|")
aTitle = Split(sTitle,"|")
aTabID = Split(sTabID,"|")
aModID = Split(sModID,"|")
aTypID = Split(sTypID,"|")
aUrlLI = Split(sUrlLI,"|")
aUrlVW = Split(sUrlVW,"|")

Function fChkRID(xID)
Dim nID : nID=-1
  For i=0 To uBound(aRssID)
  If xID=aRssID(i) Then
    nID = i
	Exit For
  End If
  Next
fChkRID = nID
End Function

Function fFilRss(xStr)
   xText = xStr&""
   xText = Replace(xText, chr(34), "&#34;")
   xText = Replace(xText, "'", "&#39;") 
   xText = Replace(xText,"<","&lt;")
   xText = Replace(xText,">","&gt;")  
   fFilRss = xText
End Function

Function fToGMT(sDate)
  Dim dWeek,dMonth
  Dim strZero,strZone
  strZero="00"
  strZone="+0800"
  dWeek=Array("Sun","Mon","Tue","Wes","Thu","Fri","Sat")
  dMonth=Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
  fToGMT = dWeek(WeekDay(sDate)-1)&", "&Right(strZero&Day(sDate),2)&" "&dMonth(Month(sDate)-1)&" "&Year(sDate)&" "&Right(strZero&Hour(sDate),2)&":"&Right(strZero&Minute(sDate),2)&":"&Right(strZero&Second(sDate),2)&" "&strZone
End Function



%>