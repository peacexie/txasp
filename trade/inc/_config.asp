
<%

verNow = "1"

Set rs=Server.CreateObject("Adodb.Recordset")
 
%>


<!--#include file="../../inc/home/func3.asp"-->
<!--#include file="../../pfile/lang/vpage.asp"-->

<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->

<!--#include file="../../upfile/sys/para/tmstyp2.asp" -->
<!--#include file="../../upfile/sys/para/tmsPara.asp" -->
<!--#include file="../../sadm/func2/func_sfile.asp" -->

<%

' xTemp = "<div class='SysI01 SysI00'><a href='($Page$)?ID=($UsrID)&KeyID=($KeyID)'>($InfSubj)</a></div>"
' xSql = "SELECT TOP 5 KeyID,InfSubj,SetSubj,LogAUser FROM ["&rTab&"] WHERE KeyMod='"&xMD&"' ORDER BY LogATime DESC"
' 得到 资料列表，新闻，图片
' ListPub(sTmp2,16,"")
Function ListInf1(xTemp,xLen,xMD,xSql)
 Dim s,s0 : s="" 
 Dim KeyID,InfSubj,LogATime,ImgName,LogAUser,rPage,rTab
 rTab="TradeInfo"
 rPage="iview.asp"
 'xSql = Replace(xSql,"(rTab)",rTab)
 rs.Open xSql,conn,1,1 
 Do While NOT rs.EOF
   KeyID = rs("KeyID")
   InfSubj = Show_SLen(rs("InfSubj"),xLen)
   SetSubj = rs("SetSubj")
   InfSubj = Show_sTitle(InfSubj,SetSubj)
   LogAUser = rs("LogAUser")
   LogATime = FormatDateTime(rs("LogATime"),2)
   s0 = xTemp
   s0 = Replace(s0,"($KeyID)",KeyID)
   s0 = Replace(s0,"($InfSubj)",InfSubj)
   s0 = Replace(s0,"($UsrID)",LogAUser)
   s0 = Replace(s0,"($Page$)",rPage)
   s0 = Replace(s0,"($LogATime)",LogATime)
   s=s&s0
 rs.MoveNext
 Loop
 rs.Close()
 ListInf1=s
End Function

' 得到 上下页
Function ListPNext(xTab,xMod,xTyp,xID,xWhr)
  Dim sql,iID,iSubj,iStr
  sql="SELECT TOP 1 KeyID,InfSubj FROM ["&xTab&"] WHERE "
  If xTyp<>"" Then
    sql=sql& " InfType LIKE '%"&xTyp&"%' "
  Else
    sql=sql& " KeyMod='"&xMod&"' "
  End If
  If xWhr=">" Then
    sql=sql& " AND KeyID>'"&ID&"' ORDER BY LogATime ASC "
  Else
    sql=sql& " AND KeyID<'"&ID&"' ORDER BY LogATime DESC "
  End If
  rs.Open sql,conn,1,1 
  if NOT rs.EOF then
    iID = rs("KeyID")
	iSubj = Show_Text(rs("InfSubj"))
	iStr = "<a href='?KeyID="&iID&"' target='_blank' >"&iSubj&"</a>"
  else
    If vMsg_InfNull&""="" Then vMsg_InfNull="(没有了)"
	iStr = "<font color=gray>"&vMsg_InfNull&"</font>"
  end if 
  rs.Close()
  ListPNext = iStr
End Function

%>


