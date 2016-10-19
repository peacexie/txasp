<!--#include file="config.asp"-->
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
KT = RequestS("KT",3,24)
TP = RequestS("TP",3,24)
T1 = RequestS("T1","D","1900-12-31")
T2 = RequestS("T2","D","2100-12-31")

If yAct="Del" then
  For iy = 1 To Request.Form("yID").Count
    iID = Request.Form("yID").item(iy)
	Call rs_DoSql(conn,"DELETE FROM [MemCard] WHERE CrdID='"&iID&"'")
  Next
  Msg = " 删除完成！" 
End If

 sqlk = ""
 if KW&"" <> "" then
  sqlk = sqlk& " AND ( "&KT&" LIKE '%"&KW&"%' ) " 
  If KT="" Then sqlK=""
 end if
 if TP&"" <> "" then
  sqlk = sqlk& " AND ( CrdType = '"&TP&"' ) " 
 end if
 if T1&"" <> "1900-12-31" then
  sqlk = sqlk& " AND ( CrdTime >= #"&T1&"# ) " 
 end if
 if T2&"" <> "2100-12-31" then
  sqlk = sqlk& " AND ( CrdTime < #"&T2&"# ) " 
 end if

'TabCols = "SELECT CrdNO,CrdNCode,MemType,CrdName,MemSex,MemMarry,MemIcon,MemBirth,MemField,CrdSpeci,MemMobile,MemEmail,MemGrade,MemEnd,MemLast,MemImg01,CONVERT(char(20),MemReg,20) AS MemReg " '/ TOP 1200
'TabWhere = " Member"&Mem_aMemb&" WHERE 1=1 "&sqlK&""
'Dim rs(3)
'Call rs_SPPage(conn,TabCols,TabWhere,Page,18,"MemReg","MemReg DESC",rs)
   sqlOrd = "CrdType,CrdName,CrdNO" 'LogATime,CrdTime DESC
   sql = "SELECT * FROM MemCard WHERE 1=1 "&sqlK&" ORDER BY "&sqlOrd '
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 18'00
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If

sPar = "&KW="&KW&"&KT="&KT&"&TP="&TP&"&T1="&T1&"&T2="&T2&"&PG="&Page&""

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="11" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>查询管理</strong><br>
            <a href="madd.asp">增加&gt;&gt;&gt;</a></td>
          <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
          <td align="left" nowrap> <span class="col00F">测试:</span> <a href="upage.asp" target="_blank">Search</a> <font color="#FF0000"><%=msg%></font></td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td align="right" bgcolor="#FFFFFF">&nbsp;&nbsp; </td>
          <form name="form1" method="post" action="?">
            <td align="right" nowrap>&nbsp;
              <input name="send" type="hidden" id="send" value="sch">
              <input name="T1" type="text" id="T1" value="<%=T1%>" size="10" maxlength="24">
              ~
              <input name="T2" type="text" id="T2" value="<%=T2%>" size="10" maxlength="24">
              <select id=TP name=TP style="width:90; ">
                <option value="">[选择类别]</option>
				<%=Get_rsOpt(conn,"SELECT TypName AS TypID,TypName FROM WebTyps WHERE TypMod='MemC124' ORDER BY TypName ",CrdType)%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12" maxlength="24">
              <select name="KT" id="KT">
                <option value="CrdName" >客户姓名</option>
                <option value="CrdNO"   >编号/身份证</option>
                <!--<option value="CrdNCode" >客户编号</option>-->
                
                <!--<option value="CrdType" >类别</option>-->
              </select>
<input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="11" align="center" bgcolor="f8f8f8"><!--%=RS_Page6(rs(2),Page,"?KW="&KW&"&TP="&TP&"",rs(1)) %-->
      <%=RS_Page(rs,Page,"?send=pag&KW="&KW&"&KT="&KT&"&TP="&TP&"&T1="&T1&"&T2="&T2&"",1)%> </td>
  </tr>
  <tr align="center">
    <td height="24" nowrap>NO</td>
    <td height="24" nowrap>编号/身份证</td>
    <td height="24" nowrap>名称/姓名</td>
    <td nowrap>性别</td>
    <td nowrap>头衔</td>
    <td nowrap>科室</td>
    <td height="24" nowrap>&nbsp;</td>
    <td nowrap>生日</td>
    <td nowrap>修改</td>
    <td nowrap>&nbsp;</td>
    <td height="24" nowrap>&nbsp;</td>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="13" nowrap></td>
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
CrdID = rs("CrdID")
CrdNO = rs("CrdNO")
CrdType = rs("CrdType")
CrdName = Left(rs("CrdName"),8)
CrdNCode = rs("CrdNCode")
CrdSpeci = rs("CrdSpeci")
If MemBirth="1900-1-1" Or MemBirth="1900-12-31" Then
  MemBirth = "[---]"
ElseIf inStr(MemBirth," ")>0 Then
  MemBirth = FormatDateTime(MemBirth,2)
End If
CrdTime = rs("CrdTime")&""
If CrdTime="1900-12-31" Then
 CrdTime="" '<font color=red>无资料</font>
End If
' ID Card -=> Birth ////////////////////////////
tDate = CrdNO
If Len(tDate)=15 Then
  tDate = Left(tDate,6)&"19"&Mid(tDate,7)&"X"
End If
If Len(tDate)=18 Then  
  tDate = Mid(tDate,7,8)
  tDate = Left(tDate,4)&"-"&Mid(tDate,5,2)&"-"&Right(tDate,2)
  'If isDate(tDate) Then
    'Call rs_DoSql(conn,"UPDATE MemCard SET CrdTime='"&tDate&"' WHERE CrdID='"&CrdID&"'")
  'End If
End If
If CrdTime="" Then
  'CrdTime=tDate
End If
' ////////////////////////////////////////////
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=(Page-1)*rs.PageSize+i%>
        <input name="yID" type="checkbox" id="yID" value="<%=CrdID%>"></td>
      <td nowrap><%=CrdNO%></td>
      <td nowrap><%=CrdName%></td>
      <td align="center" nowrap><%=CrdNCode%></td>
      <td align="center" nowrap><%=CrdSpeci%></td>
      <td align="center" nowrap><%=CrdType%></td>
      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap><%=CrdTime%></td>
      <td align="center" nowrap><a href="medit.asp?ID=<%=CrdID%><%=sPar%>">修改</a></td>
      <td nowrap>&nbsp;</td>
      <td nowrap>&nbsp;</td>
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
        <span id="yFlag" style="visibility:hidden ">N</span>全选
        <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td colspan="2" align="right" nowrap><select name="yAct" id="yAct">
          <option value="Del" >删除</option>
		  <!--
		  <option value="Clear" selected >清理会员</option>
          <option value="EditPW" >设置密码</option>
          <option value="Open" >审核会员</option>
		  <option value="Stop" >冻结会员</option>
		  <input name="CrdNCode" type="password" id="CrdNCode" size="12" maxlength="24">
		  -->
        </select></td>
      <td nowrap></td>
      <td nowrap><input type="submit" name="Submit" value="执行">      </td>
      <td nowrap>&nbsp;</td>
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
      <td colspan="13" align="right"></td>
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
