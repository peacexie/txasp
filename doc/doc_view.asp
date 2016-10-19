<!--#include file="dinc/_config.asp"-->
<%

ID = RequestS("ID",3,48)

  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open "SELECT * FROM [DocsNews] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
  
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = Show_Form(rs("InfSubj"))
InfTo = rs("InfTo")
xxxCont = rs("InfCont")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime") 

TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
If TypName="" Then
  TypName = "<font color=gray>"&MDName&"</font>"
End If
InfTyp2 = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&InfTyp2&"'")
LogAUser = rs_Val("","SELECT UsrName FROM AdmUser"&Adm_aUser&" WHERE UsrID='"&LogAUser&"'")
If LogAUser="" Then
  LogAUser = rs("LogAUser")
End If
InfFrom = InfTyp2&" | "&LogAUser &" ("&TypName&")"

If Request("Act")<>"Adm" Then
fRead = rs_Val("","SELECT KeyN FROM DocsLogs WHERE KeyID='"&ID&"' AND LogAUser='"&Session("InnID")&"'")
If fRead&""<>"" Then
  fRead = fRead+1
  Call rs_DoSql(conn,"UPDATE DocsLogs SET KeyN=KeyN+1,LogAddIP='"&Get_CIP()&"',LogAUser='"&Session("InnID")&"',LogATime='"&Now()&"' WHERE KeyID='"&ID&"' AND LogAUser='"&Session("InnID")&"'")
Else
  fRead = 1
  Call rs_DoSql(conn,"INSERT INTO DocsLogs(KeyID,KeyN,LogAddIP,LogAUser,LogATime)VALUES('"&ID&"',1,'"&Get_CIP()&"','"&Session("InnID")&"','"&Now()&"')")  
End If
Else
  fRead = 0
End If
MD = KeyMod

  End If
  rs.Close()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=InfSubj%> - 接收公文 - <%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">
  <table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1">
    <tr>
      <th height="24" colspan="2" style="font-size:16px; font-weight:bold"><%=InfSubj%></th>
    </tr>
    <tr>
      <td align="left" nowrap><span class="fnt666">来源:</span> &nbsp; <%=InfFrom%></td>
      <td align="left" nowrap><span class="fnt666">发布:</span> &nbsp; <%=LogATime%></td>
    </tr>
    <tr id="NewsContent1">
      <th height="24" colspan="2"><hr></th>
    </tr>
    <tr>
      <td colspan="2" align="left" class="cont01" id="NewsContent0"><%Call Show_sfData(ID,"/fcont.htx")%></td>
    </tr>
    <tr id="NewsContent2">
      <th height="24" colspan="2"><hr></th>
    </tr>
    <tr>
      <td align="left"><span class="fnt666">上一篇:</span>  &nbsp;<%=ListPNext(MD,LogATime,"<")%><br />
      <span class="fnt666">下一篇:</span> &nbsp;<%=ListPNext(MD,LogATime,">")%>      </td>
      <form id="SendMail" name="SendMail" target="_blank" action="/news/pub_email.asp" method="post">
        <td width="30%" align="left" nowrap><span class="fnt666">第 <span class="fnt000">(<%=fRead%>)</span> 次阅读</span><br />
        <a href="doc_remark.asp?ModID=<%=MD%>&ObjID=<%=KeyID%>&ObjSubj=<%=Server.URLEncode(InfSubj)%>"><span class="fnt666">评论:</span> &nbsp; (<%=rs_Count(conn,"DocsRemark WHERE KeyCode='"&ID&"'")%>)篇</a></td>
      </form>
    </tr>
  </table>
</div>
<%
If InfTo="(__Public__)" Then
  jsFlag="mTopRPub"
Else
  jsFlag="mTopList"
End If
%>
<!--#include file="dinc/inc_bot.asp"-->
<%

Function ListPNext(xMod,xTim,xWhr)
  Dim sql,iID,iSubj,iStr
  sql="SELECT TOP 1 KeyID,InfSubj FROM [DocsNews] "
  sql=sql& " WHERE KeyMod='"&xMod&"' AND (InfTo='(__Public__)' or InfTo LIKE '%"&Session("InnID")&";%') "
  If xWhr=">" Then
    sql=sql& " AND LogATime>"&cfgTimeC&""&xTim&""&cfgTimeC&" ORDER BY LogATime ASC "
  Else
    sql=sql& " AND LogATime<"&cfgTimeC&""&xTim&""&cfgTimeC&" ORDER BY LogATime DESC "
  End If
  rs.Open sql,conn,1,1 
  if NOT rs.EOF then
    iID = rs("KeyID")
	iSubj = Show_Text(rs("InfSubj"))
	iStr = "<a href='?ID="&iID&"'>"&iSubj&"</a>"
  else
	iStr = "<font color=gray>(没有了)</font>"
  end if 
  rs.Close()
  ListPNext = iStr
End Function

SET rs=Nothing 
%>
</body>
</html>
