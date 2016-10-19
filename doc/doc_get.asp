<!--#include file="dinc/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>接收公文 - <%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script type="text/javascript">
<!--#include file="../upfile/sys/doc/list_depart.js"-->
</script>
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">

<%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
Flag = Request("Flag")
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255) '类别
TG = RequestS("TG",3,48) '组别 
TU = RequestS("TU",3,48) '人员

'If Flag="NoRead" Then
If Flag="NoRead" Then
  jsFlag = "mTopRFlg"
  shName = "1个月内发布 --- 未读公文"
ElseIf Flag="Public" Then
  jsFlag = "mTopRPub"
  shName = "公开公文"
Else
  jsFlag = "mTopList"
  shName = "接收公文"
End If

	If KW&"" <> "" Then
	  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
	  sqlK = sqlK & " ) " 
	End If
	If TP&"" <> "" then
	  sqlK = sqlK & " AND (InfType LIKE '"&TP&"%') " 
	  shNam2 = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&TP&"'")
	End If
	If TU&"" <> "" then
	  sqlK = sqlK & " AND (LogAUser='"&TU&"') " 
	  shNam2 = rs_Val("","SELECT UsrName FROM AdmUser"&Adm_aUser&" WHERE UsrID='"&TU&"'")
	End If
	If shNam2&"" <> "" then
	  shName = shName&" &gt; "&shNam2
	End If
cID = 0
sID = ""
  yVal = RequestS("yVal","C",24)
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
  
If yAct="SetShow" Or yAct="SetHot" Or yAct="SetUBB" Or yAct="SetTop" Then
  If yAct="SetTop" Then
    yVal=RequestSafe(yVal,"N",888)
  Else
    yVal=Left(yVal,1)
  End If
  If sID<>"" Then
    sID = Replace(sID&"''",",''","")
	sql = " UPDATE DocsNews SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" 
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
	Call rs_DFile("DocsNews",sID,"")
    Msg = cID&" 条记录删除成功!"
End If
    sql = " SELECT DocsNews.* FROM [DocsNews] "
	If Flag="NoRead" Then
	  tim1 = cfgTimeC&DateAdd("d",-30,Date)&cfgTimeC
	  tim2 = cfgTimeC&DateAdd("d",1,Date)&cfgTimeC
	  sql =sql& " WHERE (LogATime BETWEEN "&tim1&" AND "&tim2&") AND (InfTo='(__Public__)' Or InfTo LIKE '%"&Session("InnID")&";%') "
	  sql =sql& " AND KeyID NOT IN(SELECT KeyID FROM DocsLogs WHERE LogAUser='"&Session("InnID")&"' ) "
	ElseIf Flag="Public" Then
	  sql =sql& " WHERE InfTo='(__Public__)' "
    Else
	  sql =sql& " WHERE InfTo LIKE '%"&Session("InnID")&";%' "&sqlK 'KeyMod='"&ModID&"'
	  sql =sql& " AND KeyID NOT IN(SELECT KeyID FROM DocsLogs WHERE LogAUser='"&Session("InnID")&"' ) "	
	End If
	sql =sql& " ORDER BY LogATime DESC "  ':Response.Write sql
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
  <%If Flag="" Then%>
  <!--#include file="dinc/inc_dep.asp"-->
  <div class="line08">&nbsp;</div>
  <%End If%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong>[<%=shName%>] </strong> </td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="TG" type="hidden" id="TG" value="<%=TG%>" />
              <input name="TU" type="hidden" id="TU" value="<%=TU%>" />
              <input name="Flag" type="hidden" id="Flag" value="<%=Flag%>" />
              <select class=form id=select name=TP >
                <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",TP,"Lay") %>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TU="&TU&"&TG="&TG&"&Flag="&Flag&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="50%" height="27" align="center">主题</td>
    <td height="27" align="center" nowrap>类别</td>
    <td align="center" nowrap>组别</td>
    <td align="center" nowrap>发布人</td>
    <td align="center" nowrap>发布时间</td>
    <td height="27" align="center" nowrap>阅读</td>
    </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right" nowrap></td>
  </tr>

    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = rs("InfSubj")
InfTo = rs("InfTo")
SetSubj = rs("SetSubj")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")

InfSubj = Show_sTitle(InfSubj,SetSubj)

TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
If TypName="" Then
  TypName = "<font color=gray>"&MDName&"</font>"
End If

InfTyp2 = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&InfTyp2&"'")

LogAUser = rs_Val("","SELECT UsrName FROM AdmUser"&Adm_aUser&" WHERE UsrID='"&LogAUser&"'")
If LogAUser="" Then
  LogAUser = rs("LogAUser")
End If

fRead = rs_Val("","SELECT KeyN FROM DocsLogs WHERE KeyID='"&KeyID&"' AND LogAUser='"&Session("InnID")&"'")
If fRead&""="" Then
fRead = "<span class=fntF0F>未读</span>"
End If

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%></td>
      <td align="left"><a href="doc_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a>     </td>
      <td nowrap><%=TypName%><%=""&TypNam2%></td>
      <td align="center" nowrap><%=InfTyp2%></td>
      <td align="center" nowrap><%=LogAUser%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=fRead%></td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="9">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="9" align="right"></td>
    </tr>

</table>
</div>
<!--#include file="dinc/inc_bot.asp"-->
<script type="text/javascript">

function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
} 

<%
If Flag="" Then
  If TG<>"" Then 
    Response.Write("mTypChang('"&TG&"');")
  Else
    Response.Write("mTypChang('All__User');")
  End If
End If
%>

</script>
</body>
</html>
