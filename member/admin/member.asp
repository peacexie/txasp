<!--#include file="config.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->
<!--#include file="mconfig.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

yAct = Request("yAct")
Page = RequestS("Page","N",1)
KW = RequestS("KW","C",96)
TP = RequestS("TP","C",96)
IP = RequestS("IP","C",96)

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
'sql = sql & " AND (MemID NOT IN(SELECT MemID FROM Fri_Memb)) "
'sql = sql & " AND (MemID NOT IN(SELECT MemID FROM Blg_Memb)) "
'sql = sql & " AND (MemID NOT IN(SELECT USID FROM Blg_Link)) "
  'cDB = rs_Count(conn,"[Member"&Mem_aMemb&"]"&sql)
  'Call rs_DoSql(conn,"DELETE FROM [Member"&Mem_aMemb&"]"&sql)
  'Response.Write sql
  Msg = cDB&" 条信息清理完成！"
  sqlK = ""
elseif yAct="EditPW" then
  MemPW = RequestS("MemPW",3,48)
      If MemPW <> "" Then
  cDB = 0
  sql = " SELECT MemID,MemPW FROM [Member"&Mem_aMemb&"] WHERE MemID IN("&sID&")"
  SET rs2  = Server.CreateObject("Adodb.Recordset") 
  rs2.Open sql,conn,1,3 
  Do While NOT rs2.EOF
    MemID = rs2("MemID")
	rs2("MemPW") = MD5_Mem(MemPW&MemID) 
	rs2.Update()
	cDB = cDB + 1
  rs2.MoveNext
  Loop
  rs2.Close
  SET rs2 = Nothing 
  Msg = cDB&" 条信息 密码修改 完成！"
      Else
  Msg = " 密码为空！未执行任何操作！"
      End If
	  'Response.Write MemPW&MemID
elseif yAct="Del" then
	'Call rs_DoSql(conn,"DELETE FROM Company  WHERE COID='"&iID&"'")
	'Call rs_DoSql(conn,"DELETE FROM BBSMemb  WHERE MemID='"&iID&"'")
	'Call rs_DoSql(conn,"DELETE FROM Fri_Memb WHERE MemID='"&iID&"'")
	'Call rs_DoSql(conn,"DELETE FROM Blg_Memb WHERE MemID='"&iID&"'")
  Call rs_DoSql(conn,"DELETE FROM [Member"&Mem_aMemb&"] WHERE MemID IN("&sID&")")
  Msg = cDB&" 条信息 删除完成！" 
elseif yAct="Open" then
   Call rs_DoSql(conn,"UPDATE [Member"&Mem_aMemb&"] SET MemFlag='Y' WHERE MemID IN("&sID&")")
   Msg = " 设置完成！" 
elseif yAct="Stop" then
   Call rs_DoSql(conn,"UPDATE [Member"&Mem_aMemb&"] SET MemFlag='N' WHERE MemID IN("&sID&")")
   Msg = " 设置完成！" 
end if

KW   = RequestS("KW",3,24)
TP   = RequestS("TP",3,24)
	sqlk = ""
	if KW&"" <> "" then
	  sqlk = sqlk& " AND ( MemID LIKE '%"&KW&"%' OR MemName LIKE '%"&KW&"%' OR MemEmail LIKE '%"&KW&"%' ) " 
	end if
	if TP <> "" then
	  sqlk = sqlk& " AND ( MemType='"&TP&"' ) " 
	end if
	if IP <> "" then
	  sqlk = sqlk& " AND ( LogAddIP LIKE '%"&IP&"%' OR LogEditIP LIKE '%"&IP&"%' ) " 
	end if

'TabCols = "SELECT MemID,MemPW,MemType,MemName,MemSex,MemMarry,MemIcon,MemBirth,MemField,MemTel,MemMobile,MemEmail,MemGrade,MemEnd,MemLast,MemImg01,CONVERT(char(20),MemReg,20) AS MemReg " '/ TOP 1200
'TabWhere = " Member"&Mem_aMemb&" WHERE 1=1 "&sqlK&""
'Dim rs(3)
'Call rs_SPPage(conn,TabCols,TabWhere,Page,18,"MemReg","MemReg DESC",rs)
   sql = "SELECT * FROM Member"&Mem_aMemb&" WHERE 1=1 "&sqlK&" ORDER BY LogATime DESC"
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 '2 '
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="13" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>会员管理</strong> || 
          <a href="mout.xls.asp?r=<%=Rnd_ID("",8)%>" target="_blank" title="请保存为xls后缀">导出Excel</a></td>
          <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
          <td align="left" nowrap><font color="#FF0000"><%=msg%></font></td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td align="right" bgcolor="#FFFFFF">&nbsp;&nbsp; </td>
          <form name="form1" method="post" action="?">
            <td align="right" nowrap>&nbsp;
              <input name="send" type="hidden" id="send" value="sch">
              关键字
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              IP
              <input name="IP" type="text" id="IP" value="<%=IP%>" size="12">
              <select name="TP" id="TP">
                <option value="" selected>[所有类别]</option>
                <%=Get_SOpt(mCfgCode,mCfgName,Show_Text(Request("MemType")),"")%>
              </select>
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="13" align="center" bgcolor="f8f8f8"><!--%=RS_Page6(rs(2),Page,"?KW="&KW&"&TP="&TP&"",rs(1)) %-->
      <%=RS_Page(rs,Page,"?send=pag&KW="&KW,1)%> </td>
  </tr>
  <tr align="center">
    <td height="24" nowrap>NO</td>
    <td nowrap>帐号</td>
    <td height="24" nowrap>名称</td>
    <td nowrap>类别</td>
    <td height="24" nowrap>性别</td>
    <td nowrap>注册</td>
    <td nowrap>-IP</td>
    <td nowrap>登陆</td>
    <td height="24" nowrap>-IP</td>
    <td nowrap>生日</td>
    <td nowrap>审核</td>
    <td nowrap>留言</td>
    <td height="24" nowrap>修改</td>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="15" nowrap></td>
  </tr>
  <form id="fm01y" name="fm01y" action="?" method="post">
    <%

'if rs(1)>0 then
'i = 0
'rs(0).Open
'Do Until rs(0).EOF
'i = i + 1
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
MemPW = rs("MemPW")
MemQu = rs("MemQu")
MemAn = rs("MemAn")
MemType = rs("MemType")
MemType = Get_SOpt(mCfgCode,mCfgName,MemType,"Val")
MemGrade = rs("MemGrade")
MemName = Left(rs("MemName"),12)
MemSex = rs("MemSex")
MemCard = rs("MemCard")
MemBirth = rs("MemBirth")
MemEmail = rs("MemEmail")
MemTel = rs("MemTel")
MemExp = rs("MemExp")
MemFlag = rs("MemFlag")
'LogAUser = rs("LogAUser")
LogAddIP = rs("LogAddIP")
LogATime = rs("LogATime"):LogATime = Left(LogATime,Len(LogATime)-3)
LogEditIP = rs("LogEditIP")
LogETime = rs("LogETime"):LogETime = Left(LogETime,Len(LogETime)-3)
MemType = Replace(MemType,"Mod","")
if MemSex = "F" then
  MemSex = "女"&MemMarry
elseif MemSex = "M" then
  MemSex = "男"&MemMarry
else
  MemSex = "未知"
end if
If MemBirth="1900-1-1" Or MemBirth="1900-12-31" Then
  MemBirth = "[---]"
ElseIf inStr(MemBirth," ")>0 Then
  MemBirth = FormatDateTime(MemBirth,2)
End If
If MemFlag="N" Then
  MemFlag = "<font color='#FF0000'>未审</font>"
ElseIf MemFlag="Y" Then
  MemFlag = "<font color='#0000FF'>已审</font>"
End If
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
        <input name="yID" type="checkbox" id="yID" value="<%=MemID%>"></td>
      <td nowrap><a href="../adm2memb.asp?send=send&MemID=<%=MemID%>" title="<%=MemType&vbcrlf%><%=MemGrade&vbcrlf&MemEmail&":"&MemTel%>" target="_blank"><%=Left(MemID,15)%></a></td>
      <td nowrap><a href="medit.asp?ID=<%=MemID%>&Act=View" target="_blank"><%=MemName%></a></td>
      <td nowrap><%=MemType%></td>
      <td nowrap><%=LogAUser%><%=MemSex%></td>
      <td nowrap><%=LogATime%></td>
      <td nowrap><%=LogAddIP%> </td>
      <td nowrap><%=LogETime%></td>
      <td nowrap><%=LogEditIP%></td>
      <td nowrap><%=MemBirth%></td>
      <td nowrap><%=MemFlag%></td>
      <td nowrap><a href="../../smod/gbook/info_add.asp?ModID=MemB224&LnkName=<%=MemID%>" target="_blank">留言</a> </td>
      <td nowrap><a href="../../member/admin/medit.asp?ID=<%=MemID%>" target="_blank">修改</a></td>
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
      <td colspan="3" align="left" nowrap><select name="yAct" id="yAct">
        <option value="Clear" selected >清理会员</option>
        <option value="EditPW" >设置密码</option>
        <option value="Del" >删除会员</option>
        <option value="Open" >审核会员</option>
        <option value="Stop" >冻结会员</option>
      </select></td>
      <td colspan="2" nowrap><input name="MemPW" type="password" id="MemPW" size="12" maxlength="24"></td>
      <td nowrap><input type="submit" name="Submit2" value="执行"></td>
      <td nowrap>&nbsp;</td>
      <td nowrap>&nbsp;</td>
      <td nowrap>&nbsp;</td>
      <td nowrap>&nbsp;</td>
      <td nowrap>&nbsp;</td>
    </tr>
    <%
  'rs(0).Close
  rs.Close
  end if
  set rs = nothing
%>
    <tr bgcolor="#cccccc">
      <td colspan="15" align="right"></td>
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
