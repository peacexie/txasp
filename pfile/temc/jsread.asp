
<!--#include file="../../page/_config.asp"-->

<%

'Call Chk_URL()
ID = RequestS("ID",3,48)
Server.Execute(Config_Path&"upfile/"&Replace(ID,"-","/")&"/"&"fpara.js")
ModTab = rel_IDTab(ID)



sql = "SELECT KeyID,KeyMod,KeyCode,InfSubj,InfType,SetRead,LogATime FROM "&ModTab&" WHERE KeyID='"&ID&"' "
rs.Open sql,conn,1,3 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfSubj = rs("InfSubj")
InfType = rs("InfType")
SetRead = rs("SetRead")+1
LogATime = rs("LogATime")
rs("SetRead") = SetRead
rs.Update()
  else
KeyID = ""
  end if 
rs.Close()



Response.Write vbcrlf&"tms_SetRead.innerHTML='"&SetRead&"';"
SetRemark = rs_Count(conn,"GboSend WHERE KeyCode='"&ID&"'")
LnkRemark = "remark.asp?ModID="&KeyMod&"&ObjID="&ID&"&ObjSubj="&Server.URLEncode(InfSubj)&""
LnkRemark = "<a href=\'"&LnkRemark&"\'>"&SetRemark&"篇</a>"
Response.Write vbcrlf&"tms_Remark.innerHTML='"&LnkRemark&"';"



InfPrev = Show_jsStr(ListPNext(ModTab,KeyMod,"",LogATime,"<"))
InfNext = Show_jsStr(ListPNext(ModTab,KeyMod,"",LogATime,">"))
Response.Write vbcrlf&"tms_PagePrev.innerHTML='"&InfPrev&"';"
Response.Write vbcrlf&"tms_PageNext.innerHTML='"&InfNext&"';"



Response.Write vbcrlf&"document.write('<br>KeyID:'+KeyID);"
Response.Write vbcrlf&"document.write('<br>ImgName:'+ImgName);"
Response.Write vbcrlf&"document.write('<br>InfType:'+InfType);"
Response.Write vbcrlf&"document.write('<br>LogATime:'+LogATime);"



'document.write('[iframe src="/inc/home/upd_out.asp?FlgMod=Home" width="0" height="0" style="visibility:hidden;">[/iframe>');
Set rs=Nothing

%>
