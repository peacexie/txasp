<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>后台管理中心</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<!--#include file="conpub.asp"-->
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>

<%

KeyID   = RequestS("KeyID",3,48)
KeyID_Back = KeyID
'TP = RequestS("TP",3,48)	
KW = RequestS("KW",3,12)
Page = RequestS("Page","N",1)
send = Request("send")
sqlK = ""
	if KW&"" <> "" then
	  sqlK = sqlK & " AND ( BPKView.InfSubj LIKE '%"&KW&"%' "
	  sqlK = sqlK & " OR BPKView.InfCont LIKE '%"&KW&"%' "
	  sqlK = sqlK & " ) " 
	end if

If send = "del" Then

	sql = "DELETE FROM BPKView WHERE KeyID='"&KeyID&"'"
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "sDel" Then
  sID = ""
  For iy = 1 To Request.Form("yID").Count
    iID = Request.Form("yID").item(iy)
	sID = sID &"'"&RequestSafe(iID,3,96)&"'," 
  Next
    sID = sID &"''"
	sID = Replace(sID,",''","")
	sql = "DELETE FROM BPKView WHERE KeyID IN("&sID&")"
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!" ': Response.Write Request.Form("yID").Count&sql
ElseIf send = "Clear" Then
	sql = "DELETE FROM BPKView WHERE KeyMod NOT IN(SELECT KeyID FROM BPKTitle)"
	Call rs_DoSql(conn,sql)
	sql = "DELETE FROM BPKVote WHERE KeyMod NOT IN(SELECT KeyID FROM BPKTitle)"
	Call rs_DoSql(conn,sql)
	Msg = "清除成功!" ': Response.Write Request.Form("yID").Count&sql
End If

    sql = " SELECT BPKView.*,BPKTitle.InfSubj AS TitName FROM [BPKView] "
	sql =sql& " LEFT JOIN BPKTitle ON BPKTitle.KeyID=BPKView.KeyMod "
	sql =sql& " WHERE 1=1  "&sqlK
	sql =sql& " ORDER BY BPKView.KeyID DESC" 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 '2'
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>辩论观点管理 &gt;&gt;</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm02" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              关键字:
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="10">
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="8" align="right"></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%if not rs.eof then
              Response.Write RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)
              end if%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" align="center">NO</td>
    <td height="27" align="center">辩论观点</td>
    <td height="27" align="center" nowrap>辩论主题</td>
    <td align="center">立场</td>
    <td align="center">发布-IP</td>
    <td align="center" nowrap>发布人</td>
    <td height="27" align="center" nowrap>修改</td>
    <td height="27" align="center">删除</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="10" align="right"></td>
  </tr>
  <form id="fm01y" name="fm01y" action="?" method="post">
  
  <%
  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#F8F8F8"
		end if
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
If KeyFlag="Neutral" Then
  KeyFlag = "中立"
ElseIf KeyFlag="Aegis" Then
  KeyFlag = "正方"
ElseIf KeyFlag="Oppose" Then 
  KeyFlag = "反方"
End If
InfSubj = Left(rs("InfSubj"),15)
InfCont = Left(rs("InfCont"),15)
InfCont = Replace(InfCont,chr(34)," `")
'SetRead = rs("SetRead")
SetUBB = rs("SetUBB")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
InfMemb = rs("InfMemb")
TitName = Left(rs("TitName"),8)
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")

	  %>

    <tr bgcolor="<%=col%>">
      <td align="right"><%=i%>
      <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td> <a 
			  href="/pkview.asp?ID=<%=KeyMod%>" target="_blank" title="<%=InfCont%>"><%=InfSubj%></a>
      </td>
      <td nowrap><%=TitName%></td>
      <td align="center" nowrap><%=KeyFlag%></td>
      <td align="center" nowrap><%=LogATime%>~<%=LogAddIP%><br>        </td>
      <td align="center"><a 
			href="/sys/member/adm2memb.asp?MBID=<%=InfMemb%>" target="_blank"><%=InfMemb%></a>-<a 
			href="/sys/member/madmin/ma_edit.asp?MBID=<%=InfMemb%>" target="_blank">权限</a></td>
      <td align="center" nowrap>        <a href="#info_edit.asp?KeyID=<%=KeyID%>&TP=<%=TP%>&KW=<%=KW%>&Page=<%=Page%>">编辑</a> </td>
      <td align="center"><input type="button" name="Button" value="删除" 
			  onClick="Del_YN('?send=del&KeyID=<%=KeyID%>&KW=<%=KW%>&Page=<%=Page%>','确认删除?')">
      </td>
    </tr>

  <%
  rs.movenext
  if rs.eof then exit for
  next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap><span id="yFlag" style="visibility:hidden ">N</span>全选
        <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td nowrap><input name="Page" type="hidden" id="Page" value="<%=Page%>">
      <input name="KW" type="hidden" id="KW" value="<%=KW%>">
      </td>
      <td nowrap>&nbsp;</td>
      <td nowrap>&nbsp;</td>
      <td nowrap><select name="send" id="send">
        <option value="sDel">删除</option>
        <option value="Clear">清除</option>
      </select></td>
      <td nowrap><input type="submit" name="Submit" value="执行"></td>

      <td align="center" nowrap>&nbsp;</td>
      <td align="center" nowrap>&nbsp;</td>
    </tr>
</form>
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
<%
  else
  %>
  
  <tr align="center" bgcolor="#f4f4f4">
    <td colspan="10">无观点</td>
  </tr>
  <%
  end if
	  
	  rs.close()
	  set rs = nothing
	  
	  %>
  <tr bgcolor="#999999">
    <td colspan="8" align="right"></td>
  </tr>
</table>
</body>
</html>
