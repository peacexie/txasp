<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
<style type="text/css">
td.f09 {
	font-size:12px;
	line-height:100%;
	letter-spacing:0px;
	word-wrap:break-word;
	word-break:break-all;
}
</style>
</head>
<body>
<%

yAct = Request("yAct")
Page = RequestS("Page","N",1)
KW = RequestS("KW","C",96)
KT = RequestS("KT","C",96)

  sID = ""
  For iy = 1 To Request.Form("yID").Count
    iID = Request.Form("yID").item(iy)
	sID = sID &"'"&RequestSafe(iID,3,128)&"'," 
  Next
    sID = sID &"''"
	sID = Replace(sID,",''","")

if yAct="Clear" then
      sql = " WHERE 1=1 "
sql = sql & " AND (LogATime<'"&DateAdd("m",-5,Now())&"') "
sql = sql & " AND (LogETime<'"&DateAdd("m",-5,Now())&"') "
sql = sql & " AND (MemID NOT IN(SELECT COID FROM Company)) "
sql = sql & " AND (MemID NOT IN(SELECT MemID FROM BBSMemb)) "
  'cDB = rs_Count(conn,"[SmsMember]"&sql)
  'Call rs_DoSql(conn,"DELETE FROM [SmsMember]"&sql)
  'Response.Write sql
  Msg = cDB&" 条信息清理完成！"
  sqlK = ""
elseif yAct="Del" then
  Call rs_DoSql(conn,"DELETE FROM [SmsMember] WHERE MemID IN("&sID&")")
  Msg = cDB&" 条信息 删除完成！" 
elseif yAct="Open" then
   Call rs_DoSql(conn,"UPDATE [SmsMember] SET MemFlag='Y' WHERE MemID IN("&sID&")")
   Msg = " 设置完成！" 
elseif yAct="Stop" then
   Call rs_DoSql(conn,"UPDATE [SmsMember] SET MemFlag='N' WHERE MemID IN("&sID&")")
   Msg = " 设置完成！" 
end if

'ID;Name;Mobile;IP;User","帐号;名称;手机
sqlK = ""
if KW&"" <> "" then
  If KT="ID" Then
    sqlK = sqlK& " AND ( MemID LIKE '%"&KW&"%' ) " 
  ElseIf KT="Name" Then
    sqlK = sqlK& " AND ( MemName LIKE '%"&KW&"%' ) "
  ElseIf KT="Mobile" Then
    sqlK = sqlK& " AND ( MemMobile LIKE '%"&KW&"%' ) "
  ElseIf KT="IP" Then
    sqlK = sqlK& " AND ( LogAddIP LIKE '%"&KW&"%' ) "
  ElseIf KT="User" Then
    sqlK = sqlK& " AND ( LogAUser LIKE '%"&KW&"%' ) "
  Else
    sqlK = sqlK& " AND ( MemID LIKE '%"&KW&"%' ) " 
  End If
end if

'TabCols = "SELECT MemID,MemPW,MemType,MemName,MemSex,MemMarry,MemIcon,MemBirth,MemField,MemTel,MemMobile,MemEmail,MemGrade,MemEnd,MemLast,MemImg01,CONVERT(char(20),MemReg,20) AS MemReg " '/ TOP 1200
'TabWhere = " SmsMember WHERE 1=1 "&sqlK&""
'Dim rs(3)
'Call rs_SPPage(conn,TabCols,TabWhere,Page,18,"MemReg","MemReg DESC",rs)
   sql = "SELECT * FROM SmsMember WHERE 1=1 "&sqlK&" ORDER BY LogATime DESC"
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 '2 '
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <form action="?" method="post" name="fsch9">
        <tr bgcolor="#FFFFFF">
          <td width="20%" align="center" nowrap bgcolor="#FFFFFF"><strong>短信会员</strong> | <a href="madd.asp">会员增加&gt;&gt;&gt;</a></td>
          <td align="left" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=MSG%></font></td>
          <td width="30%" align="right" nowrap bgcolor="#FFFFFF"><select name="KT" id="KT">
            <%=Get_SOpt("ID;Name;Mobile;IP;User","帐号;名称;手机;IP(IP);User(帐号)",KT,"")%>
          </select>
            <input name="KW" type="text" id="KW" value="<%=KW%>" size="8" maxlength="12">
            <input type="submit" name="Submit3" value="查询"></td>
        </tr>
      </form>
    </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="10" align="center" bgcolor="f8f8f8">
      <%=RS_Page(rs,Page,"?send=pag&KW="&KW&"&KT="&KT&"",1)%> </td>
  </tr>
  <tr align="center">
    <td height="24" nowrap>NO</td>
    <td nowrap>帐号</td>
    <td height="24" nowrap>名称</td>
    <td nowrap>手机</td>
    <td height="24" nowrap>余额</td>
    <td width="30%" nowrap>SN/Url</td>
    <td height="24" nowrap>审核</td>
    <td height="24" nowrap>修改</td>
    <td width="15%" nowrap>Create</td>
    <td width="15%" nowrap>Modify</td>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="12" nowrap></td>
  </tr>
  <form id="fm01y" name="fm01y" action="?" method="post">
    <%

  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f8f8f8"
		end if
'MemID = rs(0)("MemID")
MemID = rs("MemID")
MemMod = rs("MemMod")
MemCode = rs("MemCode")
MemName = rs("MemName")
MemMobile = rs("MemMobile")
MemBalance = rs("MemBalance")
MemFlag = rs("MemFlag")
MemUrl = rs("MemUrl")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
MemFlag = Get_State(MemFlag,"N;Y;-","未审;已审;未知")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
        <input name="yID" type="checkbox" id="yID" value="<%=MemID%>"></td>
      <td nowrap><%=Left(MemID,15)%></td>
      <td nowrap><%=MemName%></td>
      <td nowrap><%=MemMobile%></td>
      <td align="right" nowrap><%=MemBalance%>&nbsp;</td>
      <td class="f09"><%=MemCode%><br><%=MemUrl%></td>
      <td align="center" nowrap><%=MemFlag%></td>
      <td align="center" nowrap><a href="medit.asp?ID=<%=MemID%>&KW=<%=KW%>&KT=<%=KT%>&PG=<%=Page%>">修改</a></td>
      <td align="center" nowrap class="f09"><%=LogATime&"<br>"&LogAddIP&"("&LogAUser&")"%></td>
      <td align="center" nowrap class="f09"><%=LogETime&"<br>"&LogEditIP&"("&LogEUser&")"%></td>
    </tr>
    <%
    'rs(0).MoveNext()
  'Loop
  rs.movenext
  if rs.eof then exit for
  next
  %>
    <tr align="center" bgcolor="E0E0E0">
      <td height="21" align="right" nowrap><input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <span id="yFlag" style="visibility:hidden ">N</span>        <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td align="left" nowrap>全选</td>
      <td colspan="3" align="left" nowrap>&nbsp;</td>
      <td align="right" nowrap><select name="yAct" id="yAct">
        <option value="Clear" selected >清理会员</option>
        <option value="Del" >删除会员</option>
        <option value="Open" >审核会员</option>
        <option value="Stop" >冻结会员</option>
      </select></td>
      <td colspan="4" align="left" nowrap><input type="submit" name="Submit2" value="执行"></td>
    </tr>
    <%
  'rs(0).Close
  rs.Close
  end if
  set rs = nothing
%>
    <tr bgcolor="#cccccc">
      <td colspan="12" align="right"></td>
    </tr>
  </form>
</table>
<script type="text/javascript">
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.fm01y.yID.length;i++)
   {document.fm01y.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.fm01y.yID.length;i++)
   {document.fm01y.yID.item(i).checked=false;}
   }
}  
</script>
</body>
</html>
