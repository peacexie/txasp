<!--#include file="dinc/_config.asp"-->
<%				

Act = Request("Act")  'xieys&Act=AdmLogin
If Act="AdmLogin" Then
    Session("InnID") = RequestS("InnID","C",120)
	Session("InnPerm") = "{(MemFEditor);(MemFUpload)}"&rs_Val("","SELECT UsrPerm FROM AdmUser"&Adm_aUser&" WHERE UsrID='"&Session("InnID")&"'")
	Session("UsrLTime") = "From:("&Session("UsrID")&")"
   ID = Request("ID")
   Dir = Request("Dir")
   If Dir<>"" Then 
     Response.Redirect Dir&"?ID="&ID&""
   End If
Else
  Call Chk_Perm3(ModID,Config_Path&"doc/")
End If

USNM  = rs_val("","SELECT UsrName FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&Session("InnID")&"'")
UStr  = "["&Session("InnID")&"] "&USNM
nPub = rs_Count(conn,"DocsNews WHERE LogAUser='"&Session("InnID")&"'") 
nGet = rs_Count(conn,"DocsNews WHERE InfTo LIKE '%"&Session("InnID")&";%'")
tim1 = cfgTimeC&DateAdd("d",-30,Date)&cfgTimeC
tim2 = cfgTimeC&DateAdd("d",1,Date)&cfgTimeC
sqlw = " WHERE (LogATime BETWEEN "&tim1&" AND "&tim2&") AND (InfTo='(__Public__)' Or InfTo LIKE '%"&Session("InnID")&";%') "
sqlw =sqlw& " AND KeyID NOT IN(SELECT KeyID FROM DocsLogs WHERE LogAUser='"&Session("InnID")&"' ) "
nNor = rs_Count(conn,"DocsNews "&sqlw)

sql = " SELECT DocsNews.* FROM [DocsNews] "
sql =sql& " WHERE InfTo LIKE '%"&Session("InnID")&";%' "&sqlK 'KeyMod='"&ModID&"'
If Flag="NoRead" Then
  sql =sql& " AND KeyID NOT IN(SELECT KeyID FROM DocsLogs WHERE LogAUser='"&Session("InnID")&"' ) "
End If
  sql =sql& " ORDER BY LogATime DESC "  ':Response.Write sql

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script type="text/javascript">
<!--#include file="../upfile/sys/doc/list_depart.js"-->
</script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">
  <div class="line08">&nbsp;</div>
  <table width="540" border="0" align="center" cellpadding="8" cellspacing="5" bgcolor="f8f8f8">
    <tr>
      <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCFF">
          <tr align="center">
            <td height="36" colspan="2" bgcolor="#FFFFFF" style="font-size:16px; color:#009">欢迎光临！<%=sysName%></td>
          </tr>
          <tr bgcolor="999999">
            <td colspan="2" align="right"></td>
          </tr>
          <tr>
            <td width="30%" align="right" bgcolor="F8F8FF">当前用户:</td>
            <td width="70%" align="left" bgcolor="F8F8FF"><%=UStr%></td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">当前ＩＰ:</td>
            <td align="left" bgcolor="F8F8FF"><%=Get_CIP()%></td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">刷新时间:</td>
            <td align="left" bgcolor="F8F8FF"><%=Now()%></td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">上次登陆:</td>
            <td align="left" bgcolor="F8F8FF"><%=Session("UsrLTime")%></td>
          </tr>
          <tr bgcolor="999999">
            <td colspan="2" align="right"></td>
          </tr>
          <tr bgcolor="F8F8F8">
            <td colspan="2" align="left">公文统计:</td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">发布公文:</td>
            <td align="left" bgcolor="F8F8FF"><%=nPub%> 篇 <a href="info_list.asp">&gt;&gt;&gt;查看</a></td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">接收公文:</td>
            <td align="left" bgcolor="F8F8FF"><%=nGet%> 篇 <a href="doc_get.asp">&gt;&gt;&gt;查看</a></td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">未读公文:</td>
            <td align="left" bgcolor="F8F8FF"><%=nNor%> 篇 <a href="doc_get.asp?Flag=NoRead">&gt;&gt;&gt;查看</a></td>
          </tr>
          <tr bgcolor="999999">
            <td colspan="2" align="right"></td>
          </tr>
          <tr bgcolor="F8F8F8">
            <td colspan="2" align="left">其它操作:</td>
          </tr 
          >
          <tr>
            <td colspan="2" align="center" bgcolor="F8F8FF"><a href="userpw.asp">修改密码</a> | <a href="info_add.asp">发布公文</a> | <a href="../ext/login.asp?send=out">退出公文</a></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <div class="line13">&nbsp;</div>
  <table width="880" border="0" align="center" cellpadding="5" cellspacing="1" class="sysTabC1">
    <tr>
      <td height="60" align="left" valign="top" class="sysTabGB">说明：<!--<%=Session("InnPerm")%>--></td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<%jsFlag="mTopHome"%>
<!--#include file="dinc/inc_bot.asp"-->
</body>
</html>
