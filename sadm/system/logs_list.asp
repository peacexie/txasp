<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.f09 {
	font-size:12px;
	line-height:100%;
	letter-spacing:0px;
}
</style>
</head>
<body>

<%

LogTime = RequestS("LogTime",3,255)
Page = RequestS("Page","N",1)
send = Request("send") 
KW = RequestS("KW","C",48)
KT = RequestS("KT","C",48)
TM  = RequestS("TM","C",255)

TM  = Replace(TM,"/","")
TM  = Replace(TM,"-","")
TM  = Replace(TM,":","")
TM  = Replace(TM," ","")    

sqlK = ""
if KW&"" <> "" then
  If KT="IP" Then
    sqlK = sqlK& " AND ( LogIP LIKE '"&KW&"%' ) " 
  ElseIf KT="Act" Then
    sqlK = sqlK& " AND ( LogAct LIKE '%"&KW&"%' ) "
  ElseIf KT="Note" Then
    sqlK = sqlK& " AND ( LogNote LIKE '%"&KW&"%' ) "
  ElseIf KT="Page" Then
    sqlK = sqlK& " AND ( LogPag1 LIKE '%"&KW&"%' OR LogPag2 LIKE '%"&KW&"%' ) "
  ElseIf KT="Syst" Then
    sqlK = sqlK& " AND ( LogSyst LIKE '%"&KW&"%' ) "
  Else
    sqlK = sqlK& " AND ( LogUser LIKE '%"&KW&"%' ) " 
  End If
end if

if TM&"" <> "" then
  sqlK = sqlK& " AND ( LogTime<='"&TM&"%' ) " 
end if

sTime = Get_yyyymmdd(DateAdd("d",-30,Now()))

If send = "del_t12" Then
  if Len(TM)>0 then
	  sql = " DELETE FROM [AdmLogs] WHERE LogTime<='"&sTime&"%' "&sqlK
	  Call Chk_Perm1(IDPerm,"del") 
	  Call rs_DoSql(conn,sql)
	  Msg = "删除成功[T1-T2]!"
	sqlK = "" 
  end if
ElseIf send = "del_nul" Then
	  sql = " DELETE FROM [AdmLogs] WHERE LEN(LogUser)=0 "
	  Call Chk_Perm1(IDPerm,"del") 
	  Call rs_DoSql(conn,sql)
	  Msg = "删除成功[Null]!"
ElseIf send = "del_now" Then
  if Len(US&KW)>0 then
	sql = " DELETE FROM [AdmLogs] WHERE LogTime<='"&sTime&"%' "&sqlK
	Call Chk_Perm1(IDPerm,"del") 
	Call rs_DoSql(conn,sql)
	Msg = "删除成功[KW]!"
	sqlK = ""
  end if
ElseIf send = "dReset" Then
	  Call Chk_Perm1("{Admin}","del") 
	  sql = " DELETE FROM [AdmLogs] "
	  Call rs_DoSql(conn,sql)
	  Call Add_Log(conn,Session("UsrID"),"Reset Logs！","Tools","Tools")
	  Msg = "Reset删除成功!"
End If 
   sql = " SELECT * FROM AdmLogs "
   sql =sql& " WHERE 1=1 "&sqlK
   sql =sql&" ORDER BY LogTime DESC "  
   Set rs=Server.Createobject("Adodb.Recordset")      
   rs.Open Sql,conn,1,1
   rs.PageSize = 18
if int(Page)>rs.PageCount or int(Page)<0 Then
  Page = 1
End If

TX = Fmt_Time(Get_yyyymmdd(DateAdd("d",-7,Date())),"-")
'Response.Write sql


%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td rowspan="2" align="center" nowrap bgcolor="#FFFFFF"><strong>系统LOG管理</strong> <font color="#FF0000"><%=MSG%></font></td>
    <form action="?" method="post" name="fdel9">
      <td width="10%" align="right" nowrap bgcolor="#FFFFFF">
        <input name="send" type="hidden" id="send" value="del_t12">
        <a href="?send=del_now&KT=<%=KT%>&KW=<%=KW%>&TM=<%=TM%>">删除当前</a> - <a 
		href="?send=del_nul&US=<%=US%>&KW=<%=KW%>&TM=<%=TM%>&IP=<%=IP%>%>">删除空值</a>
        <input name="TM" type="text" id="TM" value="1900-12-31" size="10" maxlength="24">
        <input type="submit" name="Submit" value="删除"></td>
    </form>
  </tr>
  <tr bgcolor="#FFFFFF">
    <form action="?" method="post" name="fsch9">
      <td width="10%" align="right" nowrap bgcolor="#FFFFFF"><select name="KT" id="KT">
          <option value="User"  <%If KT="User"  Then Response.Write("selected")%>>User(操作者)</option>
          <option value="IP"    <%If KT="IP"    Then Response.Write("selected")%>>IP(操作者IP)</option>
          <option value="Act"   <%If KT="Act"   Then Response.Write("selected")%>>Act(操作名称)</option>
          <option value="Note"  <%If KT="Note"  Then Response.Write("selected")%>>Note(用户信息)</option>
          <option value="Page"  <%If KT="Page"  Then Response.Write("selected")%>>Page(发生页面)</option>
          <option value="Syst"  <%If KT="Syst"  Then Response.Write("selected")%>>Syst(System)</option>
        </select> 
        Act:
        <input name="KW" type="text" id="KW" value="<%=KW%>" size="8" maxlength="12">
      <input type="submit" name="Submit" value="查询"></td>
    </form>
  </tr>
</table>
<table width="100%"  border="0" cellpadding="1" cellspacing="1">
  <tr bgcolor="f4f4f4">
    <td colspan="7">
      <%=RS_Page(rs,Page,"?send=pag&KT="&KT&"&KW="&KW&"&TM="&TM&"",1)%>
    </td>
  </tr>
  <tr bgcolor="e0e0e0">
    <td height="21" align="right">No</td>
    <td bgcolor="e0e0e0">User</td>
    <td bgcolor="e0e0e0">LogTime</td>
    <td>Action</td>
    <td>LogIP</td>
    <td>Notes</td>
  </tr>
  <tr bgcolor="909090">
    <td colspan="6" align="right"></td>
  </tr>
  <%
  'Response.Write(chr(127)&chr(128))
  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
  
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f0f0f0"
		end if

  LogTime = fmt_Time(rs("LogTime"),"f")
  LogUser  = rs("LogUser")&""
  LogIP   = rs("LogIP")
  LogAct  = rs("LogAct")
  LogNote = rs("LogNote")&""
  LogPag1 = rs("LogPag1")
  LogPag2 = rs("LogPag2")
  LogSyst = rs("LogSyst")
  %>
  <tr bgcolor="<%=col%>">
    <td align="right"><%=(Page-1)*rs.PageSize+i%></td>
    <td nowrap><%=LogUser%>&nbsp; </td>
    <td nowrap><a href="#" title="Page1:<%=LogPag1&vbcr%>Page2:<%=LogPag2&vbcr%>System:<%=LogSyst%>">&nbsp;<%=LogTime%></a> </td>
    <td nowrap><%=LogAct%> </td>
    <td nowrap><%=LogIP%> </td>
    <td class="f09"><%=LogNote%> </td>
  </tr>
  <%
  rs.movenext
  if rs.eof then exit for
  next
  %>
  <tr bgcolor="909090">
    <td colspan="6" align="right"></td>
  </tr>
  <tr bgcolor="e0e0e0">
    <td colspan="7">
      <%=RS_Page(rs,Page,"?send=pag&KT="&KT&"&KW="&KW&"&TM="&TM&"",1)%>
    </td>
  </tr>
<%
  else
%>
  <tr align="center">
    <td colspan="6">无记录</td>
  </tr>
  <tr bgcolor="909090">
    <td colspan="6" align="right"></td>
  </tr>
<%
  end if
rs.close()
set rs = nothing  
%>
</table>
</body>
</html>
